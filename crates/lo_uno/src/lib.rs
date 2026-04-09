//! UNO-like service runtime.
//!
//! This crate provides a small, thread-safe, in-process registry of services
//! that mimic the surface area of LibreOffice's UNO component context. It
//! supports:
//!
//! - typed `UnoValue` payloads (`Null`, `Bool`, `Int`, `Number`, `String`,
//!   `Bytes`, `Array`, `Object`)
//! - singletons and lazy factories
//! - an event bus with `subscribe` / `unsubscribe` / `publish`
//! - built-in services: Echo, Info, TextTransformations
//!
//! It is *not* a full UNO bridge — there is no IPC, no IDL compiler, no
//! reference counting across processes — but it gives `lo_lok` and the CLI a
//! single place to look up "services" by name and invoke them with arguments.

use std::any::Any;
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, RwLock};

use lo_core::{LoError, Result};

pub type PropertyMap = BTreeMap<String, UnoValue>;
pub type DynInterface = Arc<dyn Any + Send + Sync>;
pub type Listener = Arc<dyn Fn(&UnoEvent) + Send + Sync>;

/// Typed value used for service arguments, return values and event payloads.
#[derive(Clone, Debug, PartialEq)]
pub enum UnoValue {
    Null,
    Bool(bool),
    Int(i64),
    Number(f64),
    String(String),
    Bytes(Vec<u8>),
    Array(Vec<UnoValue>),
    Object(PropertyMap),
}

impl UnoValue {
    pub fn string(value: impl Into<String>) -> Self {
        Self::String(value.into())
    }

    pub fn as_str(&self) -> Option<&str> {
        match self {
            Self::String(value) => Some(value.as_str()),
            _ => None,
        }
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            Self::Bool(value) => Some(*value),
            _ => None,
        }
    }

    pub fn as_i64(&self) -> Option<i64> {
        match self {
            Self::Int(value) => Some(*value),
            Self::Number(value) => Some(*value as i64),
            _ => None,
        }
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Self::Int(value) => Some(*value as f64),
            Self::Number(value) => Some(*value),
            _ => None,
        }
    }

    pub fn as_object(&self) -> Option<&PropertyMap> {
        match self {
            Self::Object(value) => Some(value),
            _ => None,
        }
    }
}

impl From<&str> for UnoValue {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl From<String> for UnoValue {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<bool> for UnoValue {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<i64> for UnoValue {
    fn from(value: i64) -> Self {
        Self::Int(value)
    }
}

impl From<f64> for UnoValue {
    fn from(value: f64) -> Self {
        Self::Number(value)
    }
}

/// An optional typed handle a service can expose alongside its method API.
#[derive(Clone)]
pub struct InterfaceEntry {
    pub name: &'static str,
    pub value: DynInterface,
}

impl InterfaceEntry {
    pub fn new<T>(name: &'static str, value: Arc<T>) -> Self
    where
        T: Any + Send + Sync + 'static,
    {
        Self { name, value }
    }
}

/// Event published by the [`ComponentContext`] event bus.
#[derive(Clone, Debug, PartialEq)]
pub struct UnoEvent {
    pub source: String,
    pub topic: String,
    pub payload: UnoValue,
}

#[derive(Copy, Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ListenerId(u64);

#[derive(Default)]
pub struct EventBus {
    next_id: AtomicU64,
    listeners: RwLock<BTreeMap<ListenerId, Listener>>,
}

impl EventBus {
    pub fn subscribe(&self, listener: Listener) -> ListenerId {
        let id = ListenerId(self.next_id.fetch_add(1, Ordering::Relaxed) + 1);
        self.listeners
            .write()
            .expect("event bus lock poisoned")
            .insert(id, listener);
        id
    }

    pub fn unsubscribe(&self, listener_id: ListenerId) -> bool {
        self.listeners
            .write()
            .expect("event bus lock poisoned")
            .remove(&listener_id)
            .is_some()
    }

    pub fn publish(&self, event: UnoEvent) {
        let listeners = self
            .listeners
            .read()
            .expect("event bus lock poisoned")
            .values()
            .cloned()
            .collect::<Vec<_>>();
        for listener in listeners {
            listener(&event);
        }
    }
}

/// Trait every UNO-style service implements. Methods take `&self` rather than
/// `&mut self` so a single service instance can be shared via `Arc` from
/// multiple call sites.
pub trait UnoService: Send + Sync {
    fn service_name(&self) -> &'static str;
    fn methods(&self) -> &'static [&'static str];
    fn invoke(&self, method: &str, args: &[UnoValue], ctx: &ComponentContext) -> Result<UnoValue>;
    fn interfaces(&self) -> Vec<InterfaceEntry> {
        Vec::new()
    }
    fn properties(&self) -> PropertyMap {
        PropertyMap::new()
    }
}

/// Lazy-instantiating service factory. Useful when service construction is
/// expensive or depends on other services in the context.
pub trait ServiceFactory: Send + Sync {
    fn service_name(&self) -> &'static str;
    fn create(&self, ctx: &ComponentContext) -> Result<Arc<dyn UnoService>>;
}

struct ComponentContextInner {
    factories: RwLock<BTreeMap<String, Arc<dyn ServiceFactory>>>,
    singletons: RwLock<BTreeMap<String, Arc<dyn UnoService>>>,
    bus: EventBus,
}

/// Cloneable, thread-safe service container. Cloning shares the underlying
/// registries via an `Arc`.
#[derive(Clone)]
pub struct ComponentContext {
    inner: Arc<ComponentContextInner>,
}

impl Default for ComponentContext {
    fn default() -> Self {
        let ctx = Self {
            inner: Arc::new(ComponentContextInner {
                factories: RwLock::new(BTreeMap::new()),
                singletons: RwLock::new(BTreeMap::new()),
                bus: EventBus::default(),
            }),
        };
        ctx.register_singleton(Arc::new(EchoService));
        ctx.register_singleton(Arc::new(TextTransformsService));
        ctx.register_singleton(Arc::new(InfoService));
        ctx
    }
}

impl ComponentContext {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn register_singleton(&self, service: Arc<dyn UnoService>) {
        self.inner
            .singletons
            .write()
            .expect("singletons lock poisoned")
            .insert(service.service_name().to_string(), service);
    }

    pub fn register_factory(&self, factory: Arc<dyn ServiceFactory>) {
        self.inner
            .factories
            .write()
            .expect("factories lock poisoned")
            .insert(factory.service_name().to_string(), factory);
    }

    pub fn create_service(&self, service_name: &str) -> Result<Arc<dyn UnoService>> {
        if let Some(service) = self
            .inner
            .singletons
            .read()
            .expect("singletons lock poisoned")
            .get(service_name)
            .cloned()
        {
            return Ok(service);
        }
        let factory = self
            .inner
            .factories
            .read()
            .expect("factories lock poisoned")
            .get(service_name)
            .cloned()
            .ok_or_else(|| LoError::InvalidInput(format!("service not found: {service_name}")))?;
        factory.create(self)
    }

    pub fn invoke(&self, service_name: &str, method: &str, args: &[UnoValue]) -> Result<UnoValue> {
        let service = self.create_service(service_name)?;
        let value = service.invoke(method, args, self)?;
        self.publish(UnoEvent {
            source: service_name.to_string(),
            topic: method.to_string(),
            payload: value.clone(),
        });
        Ok(value)
    }

    pub fn list_services(&self) -> Vec<String> {
        let mut names = self
            .inner
            .singletons
            .read()
            .expect("singletons lock poisoned")
            .keys()
            .cloned()
            .collect::<Vec<_>>();
        let factory_names = self
            .inner
            .factories
            .read()
            .expect("factories lock poisoned")
            .keys()
            .cloned()
            .collect::<Vec<_>>();
        for name in factory_names {
            if !names.contains(&name) {
                names.push(name);
            }
        }
        names
    }

    pub fn list_methods(&self, service_name: &str) -> Result<&'static [&'static str]> {
        let service = self.create_service(service_name)?;
        Ok(service.methods())
    }

    pub fn subscribe(&self, listener: Listener) -> ListenerId {
        self.inner.bus.subscribe(listener)
    }

    pub fn unsubscribe(&self, id: ListenerId) -> bool {
        self.inner.bus.unsubscribe(id)
    }

    pub fn publish(&self, event: UnoEvent) {
        self.inner.bus.publish(event)
    }
}

/// Backwards-compatible alias for callers that still spell the registry the
/// old way. Internally it's just a [`ComponentContext`].
pub type ServiceRegistry = ComponentContext;

// ---- built-in services -----------------------------------------------------

pub struct EchoService;

impl UnoService for EchoService {
    fn service_name(&self) -> &'static str {
        "com.libreoffice_rs.Echo"
    }

    fn methods(&self) -> &'static [&'static str] {
        &["ping", "echo"]
    }

    fn invoke(&self, method: &str, args: &[UnoValue], _ctx: &ComponentContext) -> Result<UnoValue> {
        match method {
            "ping" => Ok(UnoValue::string("pong")),
            "echo" => Ok(args.first().cloned().unwrap_or(UnoValue::Null)),
            other => Err(LoError::InvalidInput(format!("unknown method: {other}"))),
        }
    }
}

pub struct TextTransformsService;

impl UnoService for TextTransformsService {
    fn service_name(&self) -> &'static str {
        "com.libreoffice_rs.TextTransformations"
    }

    fn methods(&self) -> &'static [&'static str] {
        &["uppercase", "lowercase", "titlecase", "reverse"]
    }

    fn invoke(&self, method: &str, args: &[UnoValue], _ctx: &ComponentContext) -> Result<UnoValue> {
        let input = args.first().and_then(UnoValue::as_str).unwrap_or_default();
        let output = match method {
            "uppercase" => input.to_uppercase(),
            "lowercase" => input.to_lowercase(),
            "titlecase" => title_case(input),
            "reverse" => input.chars().rev().collect(),
            other => return Err(LoError::InvalidInput(format!("unknown method: {other}"))),
        };
        Ok(UnoValue::String(output))
    }
}

pub struct InfoService;

impl UnoService for InfoService {
    fn service_name(&self) -> &'static str {
        "com.libreoffice_rs.Info"
    }

    fn methods(&self) -> &'static [&'static str] {
        &["services", "ping"]
    }

    fn invoke(&self, method: &str, _args: &[UnoValue], ctx: &ComponentContext) -> Result<UnoValue> {
        match method {
            "services" => Ok(UnoValue::Array(
                ctx.list_services()
                    .into_iter()
                    .map(UnoValue::String)
                    .collect(),
            )),
            "ping" => Ok(UnoValue::string("pong")),
            other => Err(LoError::InvalidInput(format!("unknown method: {other}"))),
        }
    }
}

fn title_case(input: &str) -> String {
    let mut out = String::new();
    let mut next_upper = true;
    for ch in input.chars() {
        if ch.is_whitespace() {
            next_upper = true;
            out.push(ch);
        } else if next_upper {
            out.extend(ch.to_uppercase());
            next_upper = false;
        } else {
            out.extend(ch.to_lowercase());
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    #[test]
    fn registry_invokes_service() {
        let ctx = ComponentContext::new();
        let value = ctx
            .invoke(
                "com.libreoffice_rs.Echo",
                "echo",
                &[UnoValue::string("hello")],
            )
            .expect("invoke service");
        assert_eq!(value, UnoValue::string("hello"));
    }

    #[test]
    fn text_transformations_work() {
        let ctx = ComponentContext::new();
        let upper = ctx
            .invoke(
                "com.libreoffice_rs.TextTransformations",
                "uppercase",
                &[UnoValue::from("hello world")],
            )
            .unwrap();
        assert_eq!(upper, UnoValue::String("HELLO WORLD".to_string()));

        let title = ctx
            .invoke(
                "com.libreoffice_rs.TextTransformations",
                "titlecase",
                &[UnoValue::from("hello world")],
            )
            .unwrap();
        assert_eq!(title, UnoValue::String("Hello World".to_string()));
    }

    #[test]
    fn info_service_lists_other_services() {
        let ctx = ComponentContext::new();
        let services = ctx
            .invoke("com.libreoffice_rs.Info", "services", &[])
            .unwrap();
        let UnoValue::Array(items) = services else {
            panic!("expected array");
        };
        let names: Vec<String> = items
            .into_iter()
            .filter_map(|v| match v {
                UnoValue::String(s) => Some(s),
                _ => None,
            })
            .collect();
        assert!(names.iter().any(|n| n == "com.libreoffice_rs.Echo"));
        assert!(names
            .iter()
            .any(|n| n == "com.libreoffice_rs.TextTransformations"));
    }

    #[test]
    fn event_bus_delivers_publish_to_subscribers() {
        let ctx = ComponentContext::new();
        let received = Arc::new(Mutex::new(Vec::<String>::new()));
        let listener_received = Arc::clone(&received);
        let id = ctx.subscribe(Arc::new(move |event: &UnoEvent| {
            listener_received
                .lock()
                .unwrap()
                .push(format!("{}.{}", event.source, event.topic));
        }));
        ctx.invoke("com.libreoffice_rs.Echo", "ping", &[]).unwrap();
        assert!(ctx.unsubscribe(id));
        let log = received.lock().unwrap();
        assert!(log.iter().any(|m| m == "com.libreoffice_rs.Echo.ping"));
    }
}
