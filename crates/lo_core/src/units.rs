use std::fmt::{Display, Formatter};

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Length(pub f32);

impl Length {
    pub fn mm(value: f32) -> Self {
        Self(value)
    }

    pub fn pt(value: f32) -> Self {
        Self(value * 0.352_778)
    }

    pub fn px(value: f32) -> Self {
        Self(value * 0.264_583)
    }

    pub fn as_mm(self) -> f32 {
        self.0
    }

    pub fn as_pt(self) -> f32 {
        self.0 / 0.352_778
    }

    pub fn css(self) -> String {
        format!("{:.2}mm", self.0)
    }
}

impl Display for Length {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:.2}mm", self.0)
    }
}
