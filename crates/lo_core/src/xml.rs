fn escape_inner(value: &str, attribute: bool) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attribute => out.push_str("&quot;"),
            '\'' if attribute => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

pub fn escape_text(value: &str) -> String {
    escape_inner(value, false)
}

pub fn escape_attr(value: &str) -> String {
    escape_inner(value, true)
}

#[derive(Default, Debug, Clone)]
pub struct XmlBuilder {
    inner: String,
    stack: Vec<String>,
    indent: usize,
}

impl XmlBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn declaration(&mut self) {
        self.inner
            .push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    }

    fn write_indent(&mut self) {
        for _ in 0..self.indent {
            self.inner.push_str("  ");
        }
    }

    pub fn open(&mut self, name: &str, attrs: &[(&str, String)]) {
        self.write_indent();
        self.inner.push('<');
        self.inner.push_str(name);
        for (key, value) in attrs {
            self.inner.push(' ');
            self.inner.push_str(key);
            self.inner.push_str("=\"");
            self.inner.push_str(&escape_attr(value));
            self.inner.push('"');
        }
        self.inner.push_str(">\n");
        self.stack.push(name.to_string());
        self.indent += 1;
    }

    pub fn empty(&mut self, name: &str, attrs: &[(&str, String)]) {
        self.write_indent();
        self.inner.push('<');
        self.inner.push_str(name);
        for (key, value) in attrs {
            self.inner.push(' ');
            self.inner.push_str(key);
            self.inner.push_str("=\"");
            self.inner.push_str(&escape_attr(value));
            self.inner.push('"');
        }
        self.inner.push_str("/>\n");
    }

    pub fn text(&mut self, value: &str) {
        self.write_indent();
        self.inner.push_str(&escape_text(value));
        self.inner.push('\n');
    }

    pub fn raw(&mut self, value: &str) {
        self.write_indent();
        self.inner.push_str(value);
        if !value.ends_with('\n') {
            self.inner.push('\n');
        }
    }

    pub fn element(&mut self, name: &str, text: &str, attrs: &[(&str, String)]) {
        self.write_indent();
        self.inner.push('<');
        self.inner.push_str(name);
        for (key, value) in attrs {
            self.inner.push(' ');
            self.inner.push_str(key);
            self.inner.push_str("=\"");
            self.inner.push_str(&escape_attr(value));
            self.inner.push('"');
        }
        self.inner.push('>');
        self.inner.push_str(&escape_text(text));
        self.inner.push_str("</");
        self.inner.push_str(name);
        self.inner.push_str(">\n");
    }

    pub fn close(&mut self) {
        if let Some(name) = self.stack.pop() {
            self.indent = self.indent.saturating_sub(1);
            self.write_indent();
            self.inner.push_str("</");
            self.inner.push_str(&name);
            self.inner.push_str(">\n");
        }
    }

    pub fn finish(mut self) -> String {
        while !self.stack.is_empty() {
            self.close();
        }
        self.inner
    }
}

#[cfg(test)]
mod tests {
    use super::{escape_attr, escape_text, XmlBuilder};

    #[test]
    fn escaping_works() {
        assert_eq!(escape_text("a&b<c>"), "a&amp;b&lt;c&gt;");
        assert_eq!(escape_attr("\"quote\""), "&quot;quote&quot;");
    }

    #[test]
    fn xml_builder_writes_tags() {
        let mut xml = XmlBuilder::new();
        xml.declaration();
        xml.open("root", &[("id", "1".to_string())]);
        xml.element("child", "hello", &[]);
        let data = xml.finish();
        assert!(data.contains("<root id=\"1\">"));
        assert!(data.contains("<child>hello</child>"));
    }
}
