use crate::units::Length;

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Point {
    pub x: Length,
    pub y: Length,
}

impl Point {
    pub fn new(x: Length, y: Length) -> Self {
        Self { x, y }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Size {
    pub width: Length,
    pub height: Length,
}

impl Size {
    pub fn new(width: Length, height: Length) -> Self {
        Self { width, height }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Rect {
    pub origin: Point,
    pub size: Size,
}

impl Rect {
    pub fn new(x: Length, y: Length, width: Length, height: Length) -> Self {
        Self {
            origin: Point::new(x, y),
            size: Size::new(width, height),
        }
    }
}
