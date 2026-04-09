//! Tiny HTML escaping helper. HTML and XML escaping share the same five
//! characters; we keep a separate function so call sites read clearly.

use crate::xml::escape_text;

pub fn html_escape(value: &str) -> String {
    escape_text(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escapes_html_specials() {
        assert_eq!(html_escape("<a href=\"x\">&"), "&lt;a href=\"x\"&gt;&amp;");
    }
}
