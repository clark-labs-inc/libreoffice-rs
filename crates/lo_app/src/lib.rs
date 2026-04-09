//! Desktop application surface over the [`lo_lok`] runtime.
//!
//! This crate is the layer that turns "a registry of in-process documents"
//! into "a desktop-style office app". It owns:
//!
//! - a single `Office` shared by every window
//! - per-window state: title, kind, file path, zoom, sidebar, modified flag
//! - user preferences (theme, language, autosave interval, default zoom,
//!   sidebar visibility) persisted to `<profile>/preferences.ini`
//! - a recent-files list persisted to `<profile>/recent.tsv`
//! - a template registry with built-in Writer/Calc/Impress/Draw/Math/Base
//!   blanks plus a couple of seeded sample templates
//! - clipboard copy/paste over the [`DocumentHandle::selected_text`] hook
//! - macro recording and replay
//! - autosave snapshots to `<profile>/recovery/` and a recovery manifest
//!   that the next run can restore from
//! - HTML rendering of a "start center" landing page and per-window shells
//!   that include menubar, toolbar, sidebar and a tile preview of the doc
//!
//! It is intentionally GUI-toolkit-free: every interactive surface is
//! exposed as data (Vec<MenuSection>, Vec<ToolbarItem>, sidebar markup, …)
//! plus an `execute_window_command` entry point. A real GUI front-end can
//! consume those data structures directly, and a CLI demo can render them
//! to HTML.

use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Arc, RwLock};
use std::time::{SystemTime, UNIX_EPOCH};

use lo_core::{LoError, Result};
use lo_lok::{Callback, DocumentHandle, DocumentKind, KitEvent, LoadOptions, Office, TileRequest};
use lo_uno::{PropertyMap, UnoValue};

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct WindowId(u64);

impl Display for WindowId {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "win-{}", self.0)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Theme {
    System,
    Light,
    Dark,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SidebarPanel {
    None,
    Properties,
    Styles,
    Navigator,
    Gallery,
    DataSources,
    SlideLayouts,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Preferences {
    pub theme: Theme,
    pub language: String,
    pub autosave_interval_secs: u64,
    pub default_zoom_percent: u16,
    pub show_sidebar: bool,
}

impl Default for Preferences {
    fn default() -> Self {
        Self {
            theme: Theme::System,
            language: "en-US".to_string(),
            autosave_interval_secs: 120,
            default_zoom_percent: 100,
            show_sidebar: true,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MenuSection {
    pub label: String,
    pub items: Vec<MenuItem>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MenuItem {
    pub label: String,
    pub command: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ToolbarItem {
    pub label: String,
    pub command: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RecentDocument {
    pub title: String,
    pub kind: DocumentKind,
    pub path: String,
    pub opened_at_secs: u64,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RecoveryEntry {
    pub title: String,
    pub kind: DocumentKind,
    pub format: String,
    pub snapshot_path: String,
    pub original_path: Option<String>,
    pub saved_at_secs: u64,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TemplatePayload {
    Empty,
    Bytes { format: String, bytes: Vec<u8> },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Template {
    pub name: String,
    pub title: String,
    pub kind: DocumentKind,
    pub description: String,
    pub payload: TemplatePayload,
}

#[derive(Clone, Debug, PartialEq)]
pub struct RecordedCommand {
    pub command: String,
    pub args: PropertyMap,
}

#[derive(Clone, Debug, PartialEq)]
pub struct RecordedMacro {
    pub name: String,
    pub commands: Vec<RecordedCommand>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Clipboard {
    pub mime_type: Option<String>,
    pub text: Option<String>,
    pub bytes: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WindowInfo {
    pub id: WindowId,
    pub title: String,
    pub kind: DocumentKind,
    pub path: Option<String>,
    pub zoom_percent: u16,
    pub page_or_slide_index: usize,
    pub modified: bool,
    pub sidebar: SidebarPanel,
    pub status_text: String,
}

struct Window {
    id: WindowId,
    title: String,
    kind: DocumentKind,
    handle: DocumentHandle,
    path: Option<PathBuf>,
    zoom_percent: u16,
    page_or_slide_index: usize,
    modified: bool,
    sidebar: SidebarPanel,
    #[allow(dead_code)]
    created_at_secs: u64,
    #[allow(dead_code)]
    last_saved_at_secs: Option<u64>,
    status_text: String,
}

impl Window {
    fn info(&self) -> WindowInfo {
        WindowInfo {
            id: self.id,
            title: self.title.clone(),
            kind: self.kind,
            path: self.path.as_ref().map(|p| p.display().to_string()),
            zoom_percent: self.zoom_percent,
            page_or_slide_index: self.page_or_slide_index,
            modified: self.modified,
            sidebar: self.sidebar,
            status_text: self.status_text.clone(),
        }
    }
}

#[derive(Default)]
struct MacroRecorder {
    name: String,
    commands: Vec<RecordedCommand>,
}

struct AppState {
    next_window_id: u64,
    windows: BTreeMap<WindowId, Window>,
    active_window: Option<WindowId>,
    preferences: Preferences,
    recent_documents: Vec<RecentDocument>,
    recovery_entries: Vec<RecoveryEntry>,
    templates: BTreeMap<String, Template>,
    clipboard: Clipboard,
    event_log: Vec<String>,
    macros: BTreeMap<String, RecordedMacro>,
    recorder: Option<MacroRecorder>,
}

impl Default for AppState {
    fn default() -> Self {
        Self {
            next_window_id: 1,
            windows: BTreeMap::new(),
            active_window: None,
            preferences: Preferences::default(),
            recent_documents: Vec::new(),
            recovery_entries: Vec::new(),
            templates: default_templates(),
            clipboard: Clipboard::default(),
            event_log: Vec::new(),
            macros: BTreeMap::new(),
            recorder: None,
        }
    }
}

pub struct DesktopApp {
    office: Office,
    profile_dir: PathBuf,
    state: Arc<RwLock<AppState>>,
}

impl DesktopApp {
    pub fn new(profile_dir: impl Into<PathBuf>) -> Result<Self> {
        let profile_dir = profile_dir.into();
        fs::create_dir_all(profile_dir.join("exports"))?;
        fs::create_dir_all(profile_dir.join("recovery"))?;
        fs::create_dir_all(profile_dir.join("shells"))?;

        let office = Office::new();
        let state = Arc::new(RwLock::new(AppState::default()));
        let callback_state = Arc::clone(&state);
        let callback: Callback = Arc::new(move |event: &KitEvent| {
            let mut guard = callback_state.write().expect("desktop state lock poisoned");
            guard.event_log.push(format!("{event:?}"));
            // Keep the rolling event log capped so it stays useful as a
            // debugging aid without leaking memory in long sessions.
            if guard.event_log.len() > 500 {
                let drain = guard.event_log.len().saturating_sub(500);
                guard.event_log.drain(0..drain);
            }
        });
        office.register_callback(callback);

        let app = Self {
            office,
            profile_dir,
            state,
        };
        app.load_preferences()?;
        app.load_recent_documents()?;
        Ok(app)
    }

    pub fn office(&self) -> &Office {
        &self.office
    }

    pub fn profile_dir(&self) -> &Path {
        &self.profile_dir
    }

    pub fn preferences(&self) -> Preferences {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .preferences
            .clone()
    }

    pub fn set_preferences(&self, preferences: Preferences) {
        self.state
            .write()
            .expect("desktop state lock poisoned")
            .preferences = preferences;
    }

    pub fn save_preferences(&self) -> Result<()> {
        let preferences = self.preferences();
        let theme = match preferences.theme {
            Theme::System => "system",
            Theme::Light => "light",
            Theme::Dark => "dark",
        };
        let content = format!(
            "theme={}\nlanguage={}\nautosave_interval_secs={}\ndefault_zoom_percent={}\nshow_sidebar={}\n",
            theme,
            preferences.language,
            preferences.autosave_interval_secs,
            preferences.default_zoom_percent,
            preferences.show_sidebar
        );
        fs::write(self.profile_dir.join("preferences.ini"), content)?;
        Ok(())
    }

    pub fn load_preferences(&self) -> Result<()> {
        let path = self.profile_dir.join("preferences.ini");
        if !path.exists() {
            return Ok(());
        }
        let content = fs::read_to_string(path)?;
        let mut preferences = Preferences::default();
        for line in content.lines() {
            let (key, value) = match line.split_once('=') {
                Some(parts) => parts,
                None => continue,
            };
            match key.trim() {
                "theme" => {
                    preferences.theme = match value.trim() {
                        "light" => Theme::Light,
                        "dark" => Theme::Dark,
                        _ => Theme::System,
                    };
                }
                "language" => preferences.language = value.trim().to_string(),
                "autosave_interval_secs" => {
                    preferences.autosave_interval_secs = value.trim().parse().unwrap_or(120)
                }
                "default_zoom_percent" => {
                    preferences.default_zoom_percent = value.trim().parse().unwrap_or(100)
                }
                "show_sidebar" => {
                    preferences.show_sidebar = value.trim().eq_ignore_ascii_case("true")
                }
                _ => {}
            }
        }
        self.set_preferences(preferences);
        Ok(())
    }

    pub fn create_document(
        &self,
        kind: DocumentKind,
        title: impl Into<String>,
    ) -> Result<WindowId> {
        let title = title.into();
        let handle = self.office.open_empty(kind, title.clone())?;
        let preferences = self.preferences();
        let mut state = self.state.write().expect("desktop state lock poisoned");
        let id = WindowId(state.next_window_id);
        state.next_window_id += 1;
        let window = Window {
            id,
            title,
            kind,
            handle,
            path: None,
            zoom_percent: preferences.default_zoom_percent,
            page_or_slide_index: 0,
            modified: false,
            sidebar: if preferences.show_sidebar {
                default_sidebar(kind)
            } else {
                SidebarPanel::None
            },
            created_at_secs: now_secs(),
            last_saved_at_secs: None,
            status_text: "Ready".to_string(),
        };
        state.windows.insert(id, window);
        state.active_window = Some(id);
        Ok(id)
    }

    pub fn open_template(&self, name: &str) -> Result<WindowId> {
        let template = self
            .state
            .read()
            .expect("desktop state lock poisoned")
            .templates
            .get(name)
            .cloned()
            .ok_or_else(|| LoError::InvalidInput(format!("template not found: {name}")))?;

        match template.payload {
            TemplatePayload::Empty => self.create_document(template.kind, template.title),
            TemplatePayload::Bytes { format, bytes } => {
                self.open_bytes(template.kind, template.title, &bytes, &format, None)
            }
        }
    }

    pub fn templates(&self) -> Vec<Template> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .templates
            .values()
            .cloned()
            .collect()
    }

    pub fn open_bytes(
        &self,
        kind: DocumentKind,
        title: impl Into<String>,
        bytes: &[u8],
        format: &str,
        source_path: Option<PathBuf>,
    ) -> Result<WindowId> {
        let title = title.into();
        let handle = self.office.load_from_bytes(
            kind,
            title.clone(),
            bytes,
            LoadOptions {
                format: Some(format.to_string()),
                table_name: if kind == DocumentKind::Base {
                    Some("data".to_string())
                } else {
                    None
                },
            },
        )?;
        let preferences = self.preferences();
        let mut state = self.state.write().expect("desktop state lock poisoned");
        let id = WindowId(state.next_window_id);
        state.next_window_id += 1;
        let window = Window {
            id,
            title,
            kind,
            handle,
            path: source_path,
            zoom_percent: preferences.default_zoom_percent,
            page_or_slide_index: 0,
            modified: false,
            sidebar: if preferences.show_sidebar {
                default_sidebar(kind)
            } else {
                SidebarPanel::None
            },
            created_at_secs: now_secs(),
            last_saved_at_secs: None,
            status_text: format!("Loaded {format}"),
        };
        state.windows.insert(id, window);
        state.active_window = Some(id);
        Ok(id)
    }

    pub fn open_path(&self, path: impl Into<PathBuf>) -> Result<WindowId> {
        let path = path.into();
        let bytes = fs::read(&path)?;
        let name = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("document")
            .to_string();
        let ext = path
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_ascii_lowercase();
        let (kind, format) = infer_open_kind(&ext)?;
        self.open_bytes(kind, name, &bytes, &format, Some(path))
    }

    pub fn list_windows(&self) -> Vec<WindowInfo> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .windows
            .values()
            .map(Window::info)
            .collect()
    }

    pub fn active_window(&self) -> Option<WindowInfo> {
        let state = self.state.read().expect("desktop state lock poisoned");
        state
            .active_window
            .and_then(|id| state.windows.get(&id).map(Window::info))
    }

    pub fn set_active_window(&self, window_id: WindowId) -> Result<()> {
        let mut state = self.state.write().expect("desktop state lock poisoned");
        if state.windows.contains_key(&window_id) {
            state.active_window = Some(window_id);
            Ok(())
        } else {
            Err(LoError::InvalidInput(format!(
                "window not found: {window_id}"
            )))
        }
    }

    pub fn set_sidebar(&self, window_id: WindowId, sidebar: SidebarPanel) -> Result<()> {
        let mut state = self.state.write().expect("desktop state lock poisoned");
        let window = state
            .windows
            .get_mut(&window_id)
            .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
        window.sidebar = sidebar;
        Ok(())
    }

    pub fn set_zoom(&self, window_id: WindowId, zoom_percent: u16) -> Result<()> {
        let mut state = self.state.write().expect("desktop state lock poisoned");
        let window = state
            .windows
            .get_mut(&window_id)
            .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
        window.zoom_percent = zoom_percent.clamp(25, 400);
        Ok(())
    }

    pub fn execute_window_command(
        &self,
        window_id: WindowId,
        command: &str,
        arguments: &PropertyMap,
    ) -> Result<UnoValue> {
        // App-level commands are handled here without ever touching the
        // document handle, so they can run on a window with no selection.
        match command {
            "app:copy" => return self.copy_selection(window_id),
            "app:paste" => return self.paste_clipboard(window_id),
            "app:save" => {
                self.save_window(window_id)?;
                return Ok(UnoValue::Bool(true));
            }
            "app:save-as" => {
                let path = arguments
                    .get("path")
                    .and_then(UnoValue::as_str)
                    .ok_or_else(|| {
                        LoError::InvalidInput("app:save-as requires a path argument".to_string())
                    })?;
                self.save_window_as(window_id, path)?;
                return Ok(UnoValue::Bool(true));
            }
            "app:export" => {
                let path = arguments
                    .get("path")
                    .and_then(UnoValue::as_str)
                    .ok_or_else(|| {
                        LoError::InvalidInput("app:export requires a path argument".to_string())
                    })?;
                self.export_window(window_id, path)?;
                return Ok(UnoValue::Bool(true));
            }
            "app:autosave" => {
                self.autosave_all()?;
                return Ok(UnoValue::Bool(true));
            }
            _ => {}
        }

        let (handle, _title, kind) = {
            let state = self.state.read().expect("desktop state lock poisoned");
            let window = state
                .windows
                .get(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            (window.handle.clone(), window.title.clone(), window.kind)
        };
        let value = handle.execute_command(command, arguments)?;

        let mut state = self.state.write().expect("desktop state lock poisoned");
        {
            let window = state
                .windows
                .get_mut(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            if is_mutating_command(command) {
                window.modified = true;
                window.status_text = format!("{} updated", module_name(kind));
            } else {
                window.status_text = format!("{} executed", command);
            }
        }

        if let Some(recorder) = &mut state.recorder {
            if command.starts_with(".uno:") {
                recorder.commands.push(RecordedCommand {
                    command: command.to_string(),
                    args: arguments.clone(),
                });
            }
        }
        state.active_window = Some(window_id);
        Ok(value)
    }

    pub fn copy_selection(&self, window_id: WindowId) -> Result<UnoValue> {
        let handle = self.window_handle(window_id)?;
        let selected = handle.selected_text()?.unwrap_or_default();
        let mut state = self.state.write().expect("desktop state lock poisoned");
        state.clipboard = Clipboard {
            mime_type: Some("text/plain;charset=utf-8".to_string()),
            text: Some(selected.clone()),
            bytes: Some(selected.clone().into_bytes()),
        };
        if let Some(window) = state.windows.get_mut(&window_id) {
            window.status_text = "Selection copied".to_string();
        }
        Ok(UnoValue::String(selected))
    }

    pub fn paste_clipboard(&self, window_id: WindowId) -> Result<UnoValue> {
        let text = self
            .state
            .read()
            .expect("desktop state lock poisoned")
            .clipboard
            .text
            .clone()
            .unwrap_or_default();
        let mut args = PropertyMap::new();
        args.insert("text".to_string(), UnoValue::String(text));
        self.execute_window_command(window_id, ".uno:InsertText", &args)
    }

    pub fn start_macro_recording(&self, name: impl Into<String>) {
        self.state
            .write()
            .expect("desktop state lock poisoned")
            .recorder = Some(MacroRecorder {
            name: name.into(),
            commands: Vec::new(),
        });
    }

    pub fn stop_macro_recording(&self) -> Option<RecordedMacro> {
        let mut state = self.state.write().expect("desktop state lock poisoned");
        let recorder = state.recorder.take()?;
        let recorded = RecordedMacro {
            name: recorder.name.clone(),
            commands: recorder.commands,
        };
        state.macros.insert(recorded.name.clone(), recorded.clone());
        Some(recorded)
    }

    pub fn macros(&self) -> Vec<RecordedMacro> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .macros
            .values()
            .cloned()
            .collect()
    }

    pub fn play_macro(&self, name: &str, window_id: WindowId) -> Result<()> {
        let recorded = self
            .state
            .read()
            .expect("desktop state lock poisoned")
            .macros
            .get(name)
            .cloned()
            .ok_or_else(|| LoError::InvalidInput(format!("macro not found: {name}")))?;
        for command in recorded.commands {
            self.execute_window_command(window_id, &command.command, &command.args)?;
        }
        Ok(())
    }

    pub fn save_window(&self, window_id: WindowId) -> Result<PathBuf> {
        let path = {
            let state = self.state.read().expect("desktop state lock poisoned");
            if let Some(window) = state.windows.get(&window_id) {
                window.path.clone().unwrap_or_else(|| {
                    self.profile_dir.join("exports").join(format!(
                        "{}.{}",
                        slugify(&window.title),
                        default_save_format(window.kind)
                    ))
                })
            } else {
                return Err(LoError::InvalidInput(format!(
                    "window not found: {window_id}"
                )));
            }
        };
        self.save_window_as(window_id, path)
    }

    pub fn save_window_as(&self, window_id: WindowId, path: impl Into<PathBuf>) -> Result<PathBuf> {
        let path = path.into();
        let (handle, kind, title) = {
            let state = self.state.read().expect("desktop state lock poisoned");
            let window = state
                .windows
                .get(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            (window.handle.clone(), window.kind, window.title.clone())
        };
        let format = extension_or_default(&path, kind);
        let bytes = handle.save_as(&format)?;
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(&path, bytes)?;

        let mut state = self.state.write().expect("desktop state lock poisoned");
        {
            let window = state
                .windows
                .get_mut(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            window.path = Some(path.clone());
            window.modified = false;
            window.last_saved_at_secs = Some(now_secs());
            window.status_text = format!("Saved {}", path.display());
        }
        push_recent(
            &mut state.recent_documents,
            RecentDocument {
                title,
                kind,
                path: path.display().to_string(),
                opened_at_secs: now_secs(),
            },
        );
        drop(state);
        self.save_recent_documents()?;
        Ok(path)
    }

    pub fn export_window(&self, window_id: WindowId, path: impl Into<PathBuf>) -> Result<PathBuf> {
        let path = path.into();
        let (handle, kind) = {
            let state = self.state.read().expect("desktop state lock poisoned");
            let window = state
                .windows
                .get(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            (window.handle.clone(), window.kind)
        };
        let format = extension_or_default(&path, kind);
        let bytes = handle.save_as(&format)?;
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(&path, bytes)?;
        Ok(path)
    }

    pub fn close_window(&self, window_id: WindowId) -> Result<()> {
        let mut state = self.state.write().expect("desktop state lock poisoned");
        if state.windows.remove(&window_id).is_none() {
            return Err(LoError::InvalidInput(format!(
                "window not found: {window_id}"
            )));
        }
        if state.active_window == Some(window_id) {
            state.active_window = state.windows.keys().next().copied();
        }
        Ok(())
    }

    pub fn recent_documents(&self) -> Vec<RecentDocument> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .recent_documents
            .clone()
    }

    pub fn recovery_entries(&self) -> Vec<RecoveryEntry> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .recovery_entries
            .clone()
    }

    pub fn event_log(&self) -> Vec<String> {
        self.state
            .read()
            .expect("desktop state lock poisoned")
            .event_log
            .clone()
    }

    pub fn save_recent_documents(&self) -> Result<()> {
        let recent = self.recent_documents();
        let mut body = String::new();
        for entry in recent {
            body.push_str(&format!(
                "{}\t{}\t{}\t{}\n",
                module_name(entry.kind),
                sanitize_field(&entry.title),
                sanitize_field(&entry.path),
                entry.opened_at_secs
            ));
        }
        fs::write(self.profile_dir.join("recent.tsv"), body)?;
        Ok(())
    }

    pub fn load_recent_documents(&self) -> Result<()> {
        let path = self.profile_dir.join("recent.tsv");
        if !path.exists() {
            return Ok(());
        }
        let body = fs::read_to_string(path)?;
        let mut recent = Vec::new();
        for line in body.lines() {
            let parts = line.split('\t').collect::<Vec<_>>();
            if parts.len() != 4 {
                continue;
            }
            let kind = match kind_from_module_name(parts[0]) {
                Some(kind) => kind,
                None => continue,
            };
            recent.push(RecentDocument {
                kind,
                title: parts[1].to_string(),
                path: parts[2].to_string(),
                opened_at_secs: parts[3].parse().unwrap_or(0),
            });
        }
        self.state
            .write()
            .expect("desktop state lock poisoned")
            .recent_documents = recent;
        Ok(())
    }

    pub fn autosave_all(&self) -> Result<Vec<RecoveryEntry>> {
        let windows = {
            let state = self.state.read().expect("desktop state lock poisoned");
            state
                .windows
                .values()
                .map(|window| {
                    (
                        window.id,
                        window.title.clone(),
                        window.kind,
                        window.handle.clone(),
                        window.path.as_ref().map(|p| p.display().to_string()),
                    )
                })
                .collect::<Vec<_>>()
        };

        let mut entries = Vec::new();
        for (id, title, kind, handle, original_path) in windows {
            let format = autosave_format(kind).to_string();
            let bytes = handle.save_as(&format)?;
            let snapshot_path = self.profile_dir.join("recovery").join(format!(
                "{}-{}.{}",
                id,
                slugify(&title),
                format
            ));
            fs::write(&snapshot_path, bytes)?;
            entries.push(RecoveryEntry {
                title,
                kind,
                format,
                snapshot_path: snapshot_path.display().to_string(),
                original_path,
                saved_at_secs: now_secs(),
            });
        }

        self.state
            .write()
            .expect("desktop state lock poisoned")
            .recovery_entries = entries.clone();
        self.save_recovery_manifest()?;
        Ok(entries)
    }

    pub fn save_recovery_manifest(&self) -> Result<()> {
        let entries = self.recovery_entries();
        let mut body = String::new();
        for entry in entries {
            body.push_str(&format!(
                "{}\t{}\t{}\t{}\t{}\t{}\n",
                module_name(entry.kind),
                sanitize_field(&entry.title),
                sanitize_field(&entry.format),
                sanitize_field(&entry.snapshot_path),
                sanitize_field(entry.original_path.as_deref().unwrap_or("")),
                entry.saved_at_secs,
            ));
        }
        fs::write(self.profile_dir.join("recovery.tsv"), body)?;
        Ok(())
    }

    pub fn recover_autosaves(&self) -> Result<Vec<WindowId>> {
        let manifest_path = self.profile_dir.join("recovery.tsv");
        if !manifest_path.exists() {
            return Ok(Vec::new());
        }
        let manifest = fs::read_to_string(manifest_path)?;
        let mut recovered = Vec::new();
        for line in manifest.lines() {
            let parts = line.split('\t').collect::<Vec<_>>();
            if parts.len() != 6 {
                continue;
            }
            let kind = match kind_from_module_name(parts[0]) {
                Some(kind) => kind,
                None => continue,
            };
            let title = format!("{} (Recovered)", parts[1]);
            let format = parts[2].to_string();
            let snapshot_path = PathBuf::from(parts[3]);
            if !snapshot_path.exists() {
                continue;
            }
            let bytes = fs::read(snapshot_path)?;
            let id = self.open_bytes(kind, title, &bytes, &format, None)?;
            recovered.push(id);
        }
        Ok(recovered)
    }

    pub fn render_start_center_html(&self) -> String {
        let state = self.state.read().expect("desktop state lock poisoned");
        let mut templates = String::new();
        for template in state.templates.values() {
            templates.push_str(&format!(
                "<div class=\"card\"><h3>{}</h3><p>{}</p><p><strong>{}</strong></p></div>",
                html_escape(&template.title),
                html_escape(&template.description),
                module_name(template.kind)
            ));
        }
        let mut recent = String::new();
        for entry in &state.recent_documents {
            recent.push_str(&format!(
                "<li><strong>{}</strong> — {} — {}</li>",
                html_escape(&entry.title),
                module_name(entry.kind),
                html_escape(&entry.path)
            ));
        }
        if recent.is_empty() {
            recent.push_str("<li>No recent documents yet.</li>");
        }
        format!(
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>LibreOffice RS Start Center</title><style>{}</style></head><body><div class=\"shell\"><header><h1>LibreOffice RS Start Center</h1><p>Pure Rust desktop shell over the native document runtime.</p></header><section><h2>Templates</h2><div class=\"grid\">{}</div></section><section><h2>Recent Files</h2><ul>{}</ul></section></div></body></html>",
            desktop_css(),
            templates,
            recent
        )
    }

    pub fn save_start_center(&self, path: impl Into<PathBuf>) -> Result<PathBuf> {
        let path = path.into();
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(&path, self.render_start_center_html())?;
        Ok(path)
    }

    pub fn render_window_html(&self, window_id: WindowId) -> Result<String> {
        let (info, handle, menus, toolbar) = {
            let state = self.state.read().expect("desktop state lock poisoned");
            let window = state
                .windows
                .get(&window_id)
                .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
            (
                window.info(),
                window.handle.clone(),
                menu_model(window.kind),
                toolbar_model(window.kind),
            )
        };
        let preview = preview_markup(&handle)?;
        let menubar = menus
            .iter()
            .map(|section| {
                let items = section
                    .items
                    .iter()
                    .map(|item| {
                        if let Some(command) = &item.command {
                            format!(
                                "<span class=\"menu-item\" data-command=\"{}\">{}</span>",
                                html_escape(command),
                                html_escape(&item.label)
                            )
                        } else {
                            format!("<span class=\"menu-item\">{}</span>", html_escape(&item.label))
                        }
                    })
                    .collect::<Vec<_>>()
                    .join("");
                format!(
                    "<div class=\"menu-section\"><strong>{}</strong><div class=\"menu-dropdown\">{}</div></div>",
                    html_escape(&section.label),
                    items
                )
            })
            .collect::<Vec<_>>()
            .join("");
        let toolbar_html = toolbar
            .iter()
            .map(|item| {
                format!(
                    "<button data-command=\"{}\">{}</button>",
                    html_escape(&item.command),
                    html_escape(&item.label)
                )
            })
            .collect::<Vec<_>>()
            .join("");
        let sidebar_html = sidebar_markup(info.kind, info.sidebar);
        Ok(format!(
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>{}</title><style>{}</style></head><body><div class=\"desktop\"><header class=\"titlebar\"><div><strong>{}</strong> — {}</div><div>{}</div></header><nav class=\"menubar\">{}</nav><div class=\"toolbar\">{}</div><main class=\"workspace\"><aside class=\"sidebar\">{}</aside><section class=\"document\">{}</section></main><footer class=\"statusbar\"><span>{}</span><span>Zoom: {}%</span><span>{}</span></footer></div></body></html>",
            html_escape(&info.title),
            desktop_css(),
            html_escape(&info.title),
            module_name(info.kind),
            html_escape(info.path.as_deref().unwrap_or("Unsaved document")),
            menubar,
            toolbar_html,
            sidebar_html,
            preview,
            html_escape(&info.status_text),
            info.zoom_percent,
            if info.modified { "Modified" } else { "Saved" }
        ))
    }

    pub fn save_window_shell(
        &self,
        window_id: WindowId,
        path: impl Into<PathBuf>,
    ) -> Result<PathBuf> {
        let path = path.into();
        let html = self.render_window_html(window_id)?;
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(&path, html)?;
        Ok(path)
    }

    pub fn workspace_report(&self) -> String {
        let state = self.state.read().expect("desktop state lock poisoned");
        let mut out = String::new();
        out.push_str("LibreOffice RS Desktop Report\n");
        out.push_str("============================\n\n");
        out.push_str(&format!("Profile: {}\n", self.profile_dir.display()));
        out.push_str(&format!("Open windows: {}\n", state.windows.len()));
        out.push_str(&format!(
            "Recent documents: {}\n",
            state.recent_documents.len()
        ));
        out.push_str(&format!(
            "Recovery entries: {}\n",
            state.recovery_entries.len()
        ));
        out.push_str(&format!("Recorded macros: {}\n\n", state.macros.len()));
        for window in state.windows.values() {
            out.push_str(&format!(
                "- {} [{}] {}\n",
                window.title,
                module_name(window.kind),
                window
                    .path
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| "(unsaved)".to_string())
            ));
        }
        if !state.event_log.is_empty() {
            out.push_str("\nRecent events:\n");
            for line in state.event_log.iter().rev().take(10).rev() {
                out.push_str("  * ");
                out.push_str(line);
                out.push('\n');
            }
        }
        out
    }

    fn window_handle(&self, window_id: WindowId) -> Result<DocumentHandle> {
        let state = self.state.read().expect("desktop state lock poisoned");
        let window = state
            .windows
            .get(&window_id)
            .ok_or_else(|| LoError::InvalidInput(format!("window not found: {window_id}")))?;
        Ok(window.handle.clone())
    }
}

// ---- helpers ---------------------------------------------------------------

fn push_recent(recent: &mut Vec<RecentDocument>, document: RecentDocument) {
    recent.retain(|entry| entry.path != document.path);
    recent.insert(0, document);
    if recent.len() > 20 {
        recent.truncate(20);
    }
}

fn preview_markup(handle: &DocumentHandle) -> Result<String> {
    let tile = handle.render_tile(TileRequest {
        width: 1180,
        height: 820,
    })?;
    let markup = String::from_utf8(tile.bytes)
        .map_err(|_| LoError::InvalidInput("preview is not valid UTF-8".to_string()))?;
    Ok(markup)
}

fn is_mutating_command(command: &str) -> bool {
    !matches!(
        command,
        ".uno:EvaluateCell" | ".uno:ToMathML" | ".uno:ExecuteQuery" | ".uno:GetSelectionText"
    ) && command.starts_with(".uno:")
}

fn infer_open_kind(extension: &str) -> Result<(DocumentKind, String)> {
    match extension {
        "txt" => Ok((DocumentKind::Writer, "txt".to_string())),
        "md" | "markdown" => Ok((DocumentKind::Writer, "md".to_string())),
        "csv" => Ok((DocumentKind::Calc, "csv".to_string())),
        "math" | "mml" | "mathml" | "odf" => Ok((DocumentKind::Math, "math".to_string())),
        "dbcsv" => Ok((DocumentKind::Base, "csv".to_string())),
        other => Err(LoError::Unsupported(format!(
            "open/import surface for .{other} is not implemented in this pure-Rust workspace"
        ))),
    }
}

fn extension_or_default(path: &Path, kind: DocumentKind) -> String {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| default_save_format(kind).to_string())
}

fn default_save_format(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Writer => "odt",
        DocumentKind::Calc => "ods",
        DocumentKind::Impress => "odp",
        DocumentKind::Draw => "odg",
        // Math goes through `lo_math::save_as` which doesn't have an "odf"
        // path (that lives in `lo_odf`); use mathml as the default.
        DocumentKind::Math => "mathml",
        DocumentKind::Base => "odb",
    }
}

fn autosave_format(kind: DocumentKind) -> &'static str {
    default_save_format(kind)
}

fn default_sidebar(kind: DocumentKind) -> SidebarPanel {
    match kind {
        DocumentKind::Writer => SidebarPanel::Styles,
        DocumentKind::Calc => SidebarPanel::Properties,
        DocumentKind::Impress => SidebarPanel::SlideLayouts,
        DocumentKind::Draw => SidebarPanel::Gallery,
        DocumentKind::Math => SidebarPanel::Properties,
        DocumentKind::Base => SidebarPanel::DataSources,
    }
}

fn menu_model(kind: DocumentKind) -> Vec<MenuSection> {
    let mut sections = vec![
        MenuSection {
            label: "File".to_string(),
            items: vec![
                menu("New", None),
                menu("Open", None),
                menu("Save", Some("app:save")),
                menu("Save As", Some("app:save-as")),
                menu("Export", Some("app:export")),
                menu("Recent Files", None),
            ],
        },
        MenuSection {
            label: "Edit".to_string(),
            items: vec![
                menu("Copy", Some("app:copy")),
                menu("Paste", Some("app:paste")),
                menu("Select All", Some(".uno:SelectAll")),
                menu("Find", None),
                menu("Find & Replace", None),
            ],
        },
        MenuSection {
            label: "View".to_string(),
            items: vec![
                menu("Zoom In", None),
                menu("Zoom Out", None),
                menu("Sidebar", None),
                menu("Styles", None),
            ],
        },
        MenuSection {
            label: "Tools".to_string(),
            items: vec![
                menu("Macros", None),
                menu("AutoSave", Some("app:autosave")),
                menu("Options", None),
            ],
        },
        MenuSection {
            label: "Help".to_string(),
            items: vec![menu("Online Help", None), menu("About", None)],
        },
    ];

    let kind_section = match kind {
        DocumentKind::Writer => Some(MenuSection {
            label: "Insert".to_string(),
            items: vec![
                menu("Text", Some(".uno:InsertText")),
                menu("Heading", Some(".uno:AppendHeading")),
                menu("Table", Some(".uno:AppendTable")),
            ],
        }),
        DocumentKind::Calc => Some(MenuSection {
            label: "Sheet".to_string(),
            items: vec![
                menu("Set Cell", Some(".uno:SetCell")),
                menu("Append Row", Some(".uno:AppendRow")),
                menu("Evaluate", Some(".uno:EvaluateCell")),
            ],
        }),
        DocumentKind::Impress => Some(MenuSection {
            label: "Slide".to_string(),
            items: vec![
                menu("New Slide", Some(".uno:InsertSlide")),
                menu("Text Box", Some(".uno:InsertTextBox")),
                menu("Bullets", Some(".uno:InsertBullets")),
                menu("Shape", Some(".uno:InsertShape")),
            ],
        }),
        DocumentKind::Draw => Some(MenuSection {
            label: "Shape".to_string(),
            items: vec![
                menu("New Page", Some(".uno:InsertPage")),
                menu("Rectangle", Some(".uno:InsertShape")),
            ],
        }),
        DocumentKind::Math => Some(MenuSection {
            label: "Formula".to_string(),
            items: vec![
                menu("Set Formula", Some(".uno:SetFormula")),
                menu("MathML", Some(".uno:ToMathML")),
            ],
        }),
        DocumentKind::Base => Some(MenuSection {
            label: "Data".to_string(),
            items: vec![
                menu("Create Table", Some(".uno:CreateTable")),
                menu("Insert Row", Some(".uno:InsertRow")),
                menu("Execute Query", Some(".uno:ExecuteQuery")),
            ],
        }),
    };
    if let Some(section) = kind_section {
        sections.insert(3, section);
    }
    sections
}

fn toolbar_model(kind: DocumentKind) -> Vec<ToolbarItem> {
    let mut items = vec![
        ToolbarItem {
            label: "Save".to_string(),
            command: "app:save".to_string(),
        },
        ToolbarItem {
            label: "Copy".to_string(),
            command: "app:copy".to_string(),
        },
        ToolbarItem {
            label: "Paste".to_string(),
            command: "app:paste".to_string(),
        },
    ];
    match kind {
        DocumentKind::Writer => items.extend([
            tool("Heading", ".uno:AppendHeading"),
            tool("Text", ".uno:InsertText"),
            tool("Bold", ".uno:Bold"),
        ]),
        DocumentKind::Calc => items.extend([
            tool("Cell", ".uno:SetCell"),
            tool("Row", ".uno:AppendRow"),
            tool("Evaluate", ".uno:EvaluateCell"),
        ]),
        DocumentKind::Impress => items.extend([
            tool("Slide", ".uno:InsertSlide"),
            tool("TextBox", ".uno:InsertTextBox"),
            tool("Shape", ".uno:InsertShape"),
        ]),
        DocumentKind::Draw => items.extend([
            tool("Page", ".uno:InsertPage"),
            tool("Shape", ".uno:InsertShape"),
        ]),
        DocumentKind::Math => items.extend([
            tool("Formula", ".uno:SetFormula"),
            tool("MathML", ".uno:ToMathML"),
        ]),
        DocumentKind::Base => items.extend([
            tool("Table", ".uno:CreateTable"),
            tool("Row", ".uno:InsertRow"),
            tool("Query", ".uno:ExecuteQuery"),
        ]),
    }
    items
}

fn sidebar_markup(kind: DocumentKind, panel: SidebarPanel) -> String {
    let sections = match kind {
        DocumentKind::Writer => vec!["Styles", "Navigator", "Properties", "Gallery"],
        DocumentKind::Calc => vec!["Functions", "Styles", "Properties", "Navigator"],
        DocumentKind::Impress => vec!["Layouts", "Master Slides", "Animation", "Properties"],
        DocumentKind::Draw => vec!["Gallery", "Layers", "Properties", "Navigator"],
        DocumentKind::Math => vec!["Elements", "Properties", "Fonts"],
        DocumentKind::Base => vec!["Tables", "Queries", "Forms", "Reports"],
    };
    let title = match panel {
        SidebarPanel::None => "Hidden",
        SidebarPanel::Properties => "Properties",
        SidebarPanel::Styles => "Styles",
        SidebarPanel::Navigator => "Navigator",
        SidebarPanel::Gallery => "Gallery",
        SidebarPanel::DataSources => "Data Sources",
        SidebarPanel::SlideLayouts => "Slide Layouts",
    };
    let items = sections
        .into_iter()
        .map(|item| format!("<li>{}</li>", html_escape(item)))
        .collect::<Vec<_>>()
        .join("");
    format!("<h3>{}</h3><ul>{}</ul>", html_escape(title), items)
}

fn default_templates() -> BTreeMap<String, Template> {
    let mut templates = BTreeMap::new();
    templates.insert(
        "writer:blank".to_string(),
        Template {
            name: "writer:blank".to_string(),
            title: "Blank Writer Document".to_string(),
            kind: DocumentKind::Writer,
            description: "Empty text document ready for headings, paragraphs, and tables."
                .to_string(),
            payload: TemplatePayload::Empty,
        },
    );
    templates.insert(
        "writer:report".to_string(),
        Template {
            name: "writer:report".to_string(),
            title: "Project Report".to_string(),
            kind: DocumentKind::Writer,
            description: "Simple report outline with summary and action items.".to_string(),
            payload: TemplatePayload::Bytes {
                format: "md".to_string(),
                bytes: b"# Executive Summary\n\n- Goal\n- Status\n- Risks\n\n## Next Steps\n\nDescribe the work here.\n".to_vec(),
            },
        },
    );
    templates.insert(
        "calc:blank".to_string(),
        Template {
            name: "calc:blank".to_string(),
            title: "Blank Spreadsheet".to_string(),
            kind: DocumentKind::Calc,
            description: "Empty workbook for data and formulas.".to_string(),
            payload: TemplatePayload::Empty,
        },
    );
    templates.insert(
        "calc:budget".to_string(),
        Template {
            name: "calc:budget".to_string(),
            title: "Budget Spreadsheet".to_string(),
            kind: DocumentKind::Calc,
            description: "Starter budget sheet in CSV form.".to_string(),
            payload: TemplatePayload::Bytes {
                format: "csv".to_string(),
                bytes: b"Category,Amount\nRevenue,1000\nExpenses,400\nNet,=B2-B3\n".to_vec(),
            },
        },
    );
    templates.insert(
        "impress:blank".to_string(),
        Template {
            name: "impress:blank".to_string(),
            title: "Blank Presentation".to_string(),
            kind: DocumentKind::Impress,
            description: "New deck for talks and status updates.".to_string(),
            payload: TemplatePayload::Empty,
        },
    );
    templates.insert(
        "draw:blank".to_string(),
        Template {
            name: "draw:blank".to_string(),
            title: "Blank Drawing".to_string(),
            kind: DocumentKind::Draw,
            description: "Empty drawing canvas for diagrams and sketches.".to_string(),
            payload: TemplatePayload::Empty,
        },
    );
    templates.insert(
        "math:blank".to_string(),
        Template {
            name: "math:blank".to_string(),
            title: "Blank Formula".to_string(),
            kind: DocumentKind::Math,
            description: "Formula editor surface for MathML and ODF formulas.".to_string(),
            payload: TemplatePayload::Bytes {
                format: "math".to_string(),
                bytes: b"x^2 + y^2 = z^2".to_vec(),
            },
        },
    );
    templates.insert(
        "base:blank".to_string(),
        Template {
            name: "base:blank".to_string(),
            title: "Blank Database".to_string(),
            kind: DocumentKind::Base,
            description: "Empty database surface with tables and query view.".to_string(),
            payload: TemplatePayload::Empty,
        },
    );
    templates
}

fn menu(label: &str, command: Option<&str>) -> MenuItem {
    MenuItem {
        label: label.to_string(),
        command: command.map(|s| s.to_string()),
    }
}

fn tool(label: &str, command: &str) -> ToolbarItem {
    ToolbarItem {
        label: label.to_string(),
        command: command.to_string(),
    }
}

fn module_name(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Writer => "Writer",
        DocumentKind::Calc => "Calc",
        DocumentKind::Impress => "Impress",
        DocumentKind::Draw => "Draw",
        DocumentKind::Math => "Math",
        DocumentKind::Base => "Base",
    }
}

fn kind_from_module_name(name: &str) -> Option<DocumentKind> {
    match name {
        "Writer" => Some(DocumentKind::Writer),
        "Calc" => Some(DocumentKind::Calc),
        "Impress" => Some(DocumentKind::Impress),
        "Draw" => Some(DocumentKind::Draw),
        "Math" => Some(DocumentKind::Math),
        "Base" => Some(DocumentKind::Base),
        _ => None,
    }
}

fn now_secs() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

fn sanitize_field(value: &str) -> String {
    value.replace('\t', " ").replace('\n', " ")
}

fn slugify(value: &str) -> String {
    let mut out = String::new();
    let mut last_dash = false;
    for ch in value.chars() {
        let lower = ch.to_ascii_lowercase();
        if lower.is_ascii_alphanumeric() {
            out.push(lower);
            last_dash = false;
        } else if !last_dash {
            out.push('-');
            last_dash = true;
        }
    }
    out.trim_matches('-').to_string()
}

fn html_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn desktop_css() -> &'static str {
    r#"
        body { margin: 0; font-family: Arial, Helvetica, sans-serif; background: #f3f5f8; color: #202326; }
        .desktop, .shell { display: flex; flex-direction: column; min-height: 100vh; }
        .titlebar { background: #2f6ee2; color: white; padding: 10px 16px; display: flex; justify-content: space-between; }
        .menubar { background: #ebeff5; border-bottom: 1px solid #d7dde7; padding: 8px 12px; display: flex; gap: 18px; flex-wrap: wrap; }
        .menu-section { position: relative; }
        .menu-dropdown { display: flex; gap: 10px; flex-wrap: wrap; font-size: 13px; margin-top: 4px; }
        .menu-item { padding: 2px 6px; background: white; border: 1px solid #d7dde7; border-radius: 4px; }
        .toolbar { display: flex; gap: 8px; padding: 10px 12px; background: #ffffff; border-bottom: 1px solid #d7dde7; flex-wrap: wrap; }
        .toolbar button { border: 1px solid #c5d0df; background: #f8fbff; border-radius: 4px; padding: 6px 10px; }
        .workspace { display: grid; grid-template-columns: 260px 1fr; gap: 16px; padding: 16px; flex: 1; }
        .sidebar { background: white; border: 1px solid #d7dde7; border-radius: 8px; padding: 16px; }
        .document { background: white; border: 1px solid #d7dde7; border-radius: 8px; padding: 8px; overflow: auto; }
        .statusbar { display: flex; justify-content: space-between; gap: 12px; background: #ebeff5; border-top: 1px solid #d7dde7; padding: 8px 12px; font-size: 13px; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; }
        .card { background: white; border: 1px solid #d7dde7; border-radius: 8px; padding: 16px; }
        ul { padding-left: 20px; }
    "#
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn templates_are_seeded() {
        let templates = default_templates();
        assert!(templates.contains_key("writer:blank"));
        assert!(templates.contains_key("calc:budget"));
    }

    #[test]
    fn slugify_is_stable() {
        assert_eq!(slugify("Quarterly Report 2026"), "quarterly-report-2026");
    }

    #[test]
    fn create_document_and_save_writes_to_profile() {
        let dir = tempdir();
        let app = DesktopApp::new(&dir).expect("desktop app");
        let id = app
            .create_document(DocumentKind::Writer, "Hello")
            .expect("create writer");
        let mut args = PropertyMap::new();
        args.insert("text".to_string(), UnoValue::String("body".to_string()));
        app.execute_window_command(id, ".uno:InsertText", &args)
            .expect("insert text");
        let path = app.save_window(id).expect("save");
        assert!(path.exists());
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn start_center_html_lists_seeded_templates() {
        let dir = tempdir();
        let app = DesktopApp::new(&dir).expect("desktop app");
        let html = app.render_start_center_html();
        assert!(html.contains("Blank Writer Document"));
        assert!(html.contains("Project Report"));
        assert!(html.contains("Budget Spreadsheet"));
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn macro_record_and_play_repeats_command() {
        let dir = tempdir();
        let app = DesktopApp::new(&dir).expect("desktop app");
        let id = app.create_document(DocumentKind::Writer, "M").unwrap();
        app.start_macro_recording("add-paragraph");
        let mut args = PropertyMap::new();
        args.insert("text".to_string(), UnoValue::String("once".to_string()));
        app.execute_window_command(id, ".uno:InsertText", &args)
            .unwrap();
        let recorded = app.stop_macro_recording().expect("recorded");
        assert_eq!(recorded.commands.len(), 1);
        // Replaying inserts the same text a second time.
        app.play_macro("add-paragraph", id).unwrap();
        let bytes = app.window_handle(id).unwrap().save_as("txt").unwrap();
        let txt = String::from_utf8(bytes).unwrap();
        assert_eq!(txt.matches("once").count(), 2);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn autosave_writes_recovery_files_and_manifest() {
        let dir = tempdir();
        let app = DesktopApp::new(&dir).expect("desktop app");
        let _writer = app.create_document(DocumentKind::Writer, "AS").unwrap();
        let _calc = app.create_document(DocumentKind::Calc, "CS").unwrap();
        let entries = app.autosave_all().expect("autosave");
        assert_eq!(entries.len(), 2);
        for entry in entries {
            assert!(std::path::Path::new(&entry.snapshot_path).exists());
        }
        assert!(dir.join("recovery.tsv").exists());
        let _ = std::fs::remove_dir_all(&dir);
    }

    fn tempdir() -> PathBuf {
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let pid = std::process::id();
        let dir = std::env::temp_dir().join(format!("lo_app_test_{pid}_{n}"));
        let _ = std::fs::remove_dir_all(&dir);
        dir
    }
}
