#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Metadata {
    pub title: String,
    pub subject: String,
    pub description: String,
    pub creator: String,
    pub created: String,
    pub modified: String,
    pub keywords: Vec<String>,
}

impl Metadata {
    pub fn titled(title: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            ..Self::default()
        }
    }
}
