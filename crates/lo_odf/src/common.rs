use std::path::Path;

use lo_core::{escape_text, Metadata, Result, XmlBuilder};
use lo_zip::{write_zip_file, ZipEntry};

pub const MIME_ODT: &str = "application/vnd.oasis.opendocument.text";
pub const MIME_ODS: &str = "application/vnd.oasis.opendocument.spreadsheet";
pub const MIME_ODP: &str = "application/vnd.oasis.opendocument.presentation";
pub const MIME_ODG: &str = "application/vnd.oasis.opendocument.graphics";
pub const MIME_ODF: &str = "application/vnd.oasis.opendocument.formula";
// LibreOffice itself writes Base documents with mimetype
// `application/vnd.oasis.opendocument.base` (confirmed against biblio.odb and
// evolocal.odb shipped with LibreOffice 26.2); `.database` is not recognised
// and causes `Error: source file could not be loaded`.
pub const MIME_ODB: &str = "application/vnd.oasis.opendocument.base";

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtraFile {
    pub path: String,
    pub media_type: String,
    pub data: Vec<u8>,
}

impl ExtraFile {
    pub fn new(
        path: impl Into<String>,
        media_type: impl Into<String>,
        data: impl Into<Vec<u8>>,
    ) -> Self {
        Self {
            path: path.into(),
            media_type: media_type.into(),
            data: data.into(),
        }
    }
}

pub fn content_root_attrs() -> Vec<(&'static str, String)> {
    vec![
        (
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0".to_string(),
        ),
        (
            "xmlns:style",
            "urn:oasis:names:tc:opendocument:xmlns:style:1.0".to_string(),
        ),
        (
            "xmlns:text",
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0".to_string(),
        ),
        (
            "xmlns:table",
            "urn:oasis:names:tc:opendocument:xmlns:table:1.0".to_string(),
        ),
        (
            "xmlns:draw",
            "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0".to_string(),
        ),
        (
            "xmlns:presentation",
            "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0".to_string(),
        ),
        ("xmlns:xlink", "http://www.w3.org/1999/xlink".to_string()),
        (
            "xmlns:fo",
            "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0".to_string(),
        ),
        (
            "xmlns:svg",
            "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0".to_string(),
        ),
        (
            "xmlns:math",
            "http://www.w3.org/1998/Math/MathML".to_string(),
        ),
        (
            "xmlns:meta",
            "urn:oasis:names:tc:opendocument:xmlns:meta:1.0".to_string(),
        ),
        ("xmlns:dc", "http://purl.org/dc/elements/1.1/".to_string()),
        (
            "xmlns:of",
            "urn:oasis:names:tc:opendocument:xmlns:of:1.2".to_string(),
        ),
        ("office:version", "1.3".to_string()),
    ]
}

fn meta_root_attrs() -> Vec<(&'static str, String)> {
    vec![
        (
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0".to_string(),
        ),
        (
            "xmlns:meta",
            "urn:oasis:names:tc:opendocument:xmlns:meta:1.0".to_string(),
        ),
        ("xmlns:dc", "http://purl.org/dc/elements/1.1/".to_string()),
        ("office:version", "1.3".to_string()),
    ]
}

fn styles_root_attrs() -> Vec<(&'static str, String)> {
    vec![
        (
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0".to_string(),
        ),
        (
            "xmlns:style",
            "urn:oasis:names:tc:opendocument:xmlns:style:1.0".to_string(),
        ),
        (
            "xmlns:text",
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0".to_string(),
        ),
        (
            "xmlns:table",
            "urn:oasis:names:tc:opendocument:xmlns:table:1.0".to_string(),
        ),
        (
            "xmlns:draw",
            "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0".to_string(),
        ),
        (
            "xmlns:fo",
            "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0".to_string(),
        ),
        (
            "xmlns:svg",
            "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0".to_string(),
        ),
        (
            "xmlns:presentation",
            "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0".to_string(),
        ),
        ("office:version", "1.3".to_string()),
    ]
}

fn settings_root_attrs() -> Vec<(&'static str, String)> {
    vec![
        (
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0".to_string(),
        ),
        ("office:version", "1.3".to_string()),
    ]
}

pub fn meta_xml(meta: &Metadata) -> String {
    let mut xml = XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-meta", &meta_root_attrs());
    xml.open("office:meta", &[]);
    xml.element(
        "meta:generator",
        concat!("libreoffice-rs/", env!("CARGO_PKG_VERSION")),
        &[],
    );
    if !meta.title.is_empty() {
        xml.element("dc:title", &meta.title, &[]);
    }
    if !meta.subject.is_empty() {
        xml.element("dc:subject", &meta.subject, &[]);
    }
    if !meta.description.is_empty() {
        xml.element("dc:description", &meta.description, &[]);
    }
    if !meta.creator.is_empty() {
        xml.element("meta:initial-creator", &meta.creator, &[]);
        xml.element("dc:creator", &meta.creator, &[]);
    }
    if !meta.created.is_empty() {
        xml.element("meta:creation-date", &meta.created, &[]);
    }
    if !meta.modified.is_empty() {
        xml.element("dc:date", &meta.modified, &[]);
    }
    if !meta.keywords.is_empty() {
        xml.element("meta:keyword", &meta.keywords.join(", "), &[]);
    }
    xml.close();
    xml.close();
    xml.finish()
}

pub fn styles_xml() -> String {
    let mut xml = XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-styles", &styles_root_attrs());
    xml.open("office:styles", &[]);
    xml.raw(
        r#"<style:style style:name="Standard" style:family="paragraph"><style:text-properties fo:font-size="11pt"/></style:style>"#,
    );
    xml.raw(
        r#"<style:style style:name="Strong" style:family="text"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold"/></style:style>"#,
    );
    xml.raw(
        r#"<style:style style:name="Emphasis" style:family="text"><style:text-properties fo:font-style="italic" style:font-style-asian="italic"/></style:style>"#,
    );
    xml.raw(
        r#"<style:style style:name="Code" style:family="text"><style:text-properties style:font-name="monospace"/></style:style>"#,
    );
    // Heading paragraph styles, referenced by lo_writer when emitting
    // text:h elements. Names use LibreOffice's `_20_` space-escape.
    for level in 1..=6 {
        xml.raw(&format!(
            "<style:style style:name=\"Heading_20_{level}\" style:display-name=\"Heading {level}\" \
style:family=\"paragraph\" style:parent-style-name=\"Standard\" style:next-style-name=\"Standard\">\
<style:text-properties fo:font-size=\"{size}pt\" fo:font-weight=\"bold\"/>\
</style:style>",
            level = level,
            size = 20 - (level - 1) * 2,
        ));
    }
    // Bullet list style shared across writer + impress.
    xml.raw(
        "<text:list-style style:name=\"L1\">\
<text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"•\">\
<style:list-level-properties text:space-before=\"6mm\" text:min-label-width=\"5mm\"/>\
</text:list-level-style-bullet>\
</text:list-style>",
    );
    xml.close();
    // Page layout + default master page so presentation/drawing documents can
    // reference draw:master-page-name="Default". Writer/Calc ignore this.
    xml.raw(
        "<office:automatic-styles>\
<style:page-layout style:name=\"PL1\">\
<style:page-layout-properties fo:page-width=\"254mm\" fo:page-height=\"190.5mm\" \
fo:margin-top=\"0mm\" fo:margin-bottom=\"0mm\" fo:margin-left=\"0mm\" fo:margin-right=\"0mm\" \
style:print-orientation=\"landscape\"/>\
</style:page-layout>\
</office:automatic-styles>",
    );
    xml.raw(
        "<office:master-styles>\
<style:master-page style:name=\"Default\" style:page-layout-name=\"PL1\"/>\
</office:master-styles>",
    );
    xml.close();
    xml.finish()
}

pub fn settings_xml() -> String {
    let mut xml = XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-settings", &settings_root_attrs());
    xml.empty("office:settings", &[]);
    xml.close();
    xml.finish()
}

pub fn manifest_xml(mimetype: &str, extras: &[ExtraFile]) -> String {
    let mut xml = XmlBuilder::new();
    xml.declaration();
    xml.open(
        "manifest:manifest",
        &[
            (
                "xmlns:manifest",
                "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0".to_string(),
            ),
            ("manifest:version", "1.3".to_string()),
        ],
    );
    xml.empty(
        "manifest:file-entry",
        &[
            ("manifest:full-path", "/".to_string()),
            ("manifest:media-type", mimetype.to_string()),
        ],
    );
    for component in ["content.xml", "styles.xml", "meta.xml", "settings.xml"] {
        xml.empty(
            "manifest:file-entry",
            &[
                ("manifest:full-path", component.to_string()),
                ("manifest:media-type", "text/xml".to_string()),
            ],
        );
    }
    for extra in extras {
        xml.empty(
            "manifest:file-entry",
            &[
                ("manifest:full-path", extra.path.clone()),
                ("manifest:media-type", extra.media_type.clone()),
            ],
        );
    }
    xml.close();
    xml.finish()
}

pub fn package_document(
    path: impl AsRef<Path>,
    mimetype: &str,
    content_xml: String,
    meta: &Metadata,
    extras: Vec<ExtraFile>,
) -> Result<()> {
    let mut entries = Vec::new();
    entries.push(ZipEntry::new("mimetype", mimetype.as_bytes().to_vec()));
    entries.push(ZipEntry::new("content.xml", content_xml.into_bytes()));
    entries.push(ZipEntry::new("styles.xml", styles_xml().into_bytes()));
    entries.push(ZipEntry::new("meta.xml", meta_xml(meta).into_bytes()));
    entries.push(ZipEntry::new("settings.xml", settings_xml().into_bytes()));
    entries.push(ZipEntry::new(
        "META-INF/manifest.xml",
        manifest_xml(mimetype, &extras).into_bytes(),
    ));
    for extra in extras {
        entries.push(ZipEntry::new(extra.path, extra.data));
    }
    write_zip_file(path, &entries)
}

pub fn image_extras(images: Vec<(String, String, Vec<u8>)>) -> Vec<ExtraFile> {
    images
        .into_iter()
        .map(|(name, media_type, data)| {
            ExtraFile::new(format!("Pictures/{name}"), media_type, data)
        })
        .collect()
}

pub fn escaped_text_paragraph(text: &str) -> String {
    format!("<text:p>{}</text:p>", escape_text(text))
}

/// Write a Base document archive.
///
/// Base documents use a leaner layout than the other ODF files: no meta.xml
/// or styles.xml, just content.xml + settings.xml + manifest. The manifest
/// only lists the parts that actually exist.
pub fn package_database_document(
    path: impl AsRef<Path>,
    content_xml: String,
    extras: Vec<ExtraFile>,
) -> Result<()> {
    let mut entries = Vec::new();
    entries.push(ZipEntry::new("mimetype", MIME_ODB.as_bytes().to_vec()));
    entries.push(ZipEntry::new("content.xml", content_xml.into_bytes()));
    entries.push(ZipEntry::new("settings.xml", settings_xml().into_bytes()));

    // Manifest that only lists parts actually in this archive.
    let mut xml = XmlBuilder::new();
    xml.declaration();
    xml.open(
        "manifest:manifest",
        &[
            (
                "xmlns:manifest",
                "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0".to_string(),
            ),
            ("manifest:version", "1.3".to_string()),
        ],
    );
    xml.empty(
        "manifest:file-entry",
        &[
            ("manifest:full-path", "/".to_string()),
            ("manifest:version", "1.3".to_string()),
            ("manifest:media-type", MIME_ODB.to_string()),
        ],
    );
    xml.empty(
        "manifest:file-entry",
        &[
            ("manifest:full-path", "content.xml".to_string()),
            ("manifest:media-type", "text/xml".to_string()),
        ],
    );
    xml.empty(
        "manifest:file-entry",
        &[
            ("manifest:full-path", "settings.xml".to_string()),
            ("manifest:media-type", "text/xml".to_string()),
        ],
    );
    for extra in &extras {
        xml.empty(
            "manifest:file-entry",
            &[
                ("manifest:full-path", extra.path.clone()),
                ("manifest:media-type", extra.media_type.clone()),
            ],
        );
    }
    xml.close();
    entries.push(ZipEntry::new(
        "META-INF/manifest.xml",
        xml.finish().into_bytes(),
    ));
    for extra in extras {
        entries.push(ZipEntry::new(extra.path, extra.data));
    }
    write_zip_file(path, &entries)
}
