use std::cmp::{max, min};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Rgba {
    pub r: u8,
    pub g: u8,
    pub b: u8,
    pub a: u8,
}

impl Rgba {
    pub const WHITE: Self = Self::rgba(255, 255, 255, 255);
    pub const BLACK: Self = Self::rgba(0, 0, 0, 255);
    pub const TRANSPARENT: Self = Self::rgba(0, 0, 0, 0);

    pub const fn rgba(r: u8, g: u8, b: u8, a: u8) -> Self {
        Self { r, g, b, a }
    }

    pub fn with_alpha(self, a: u8) -> Self {
        Self { a, ..self }
    }
}

pub fn parse_hex_color(input: &str, fallback: Rgba) -> Rgba {
    let trimmed = input.trim();
    let hex = trimmed.strip_prefix('#').unwrap_or(trimmed);
    match hex.len() {
        6 => {
            let r = u8::from_str_radix(&hex[0..2], 16).ok();
            let g = u8::from_str_radix(&hex[2..4], 16).ok();
            let b = u8::from_str_radix(&hex[4..6], 16).ok();
            match (r, g, b) {
                (Some(r), Some(g), Some(b)) => Rgba::rgba(r, g, b, 255),
                _ => fallback,
            }
        }
        8 => {
            let r = u8::from_str_radix(&hex[0..2], 16).ok();
            let g = u8::from_str_radix(&hex[2..4], 16).ok();
            let b = u8::from_str_radix(&hex[4..6], 16).ok();
            let a = u8::from_str_radix(&hex[6..8], 16).ok();
            match (r, g, b, a) {
                (Some(r), Some(g), Some(b), Some(a)) => Rgba::rgba(r, g, b, a),
                _ => fallback,
            }
        }
        3 => {
            let mut chars = hex.chars();
            let r = chars.next().and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
            let g = chars.next().and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
            let b = chars.next().and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
            match (r, g, b) {
                (Some(r), Some(g), Some(b)) => Rgba::rgba(r, g, b, 255),
                _ => fallback,
            }
        }
        _ => fallback,
    }
}

#[derive(Clone, Debug)]
pub struct RasterImage {
    pub width: u32,
    pub height: u32,
    pub pixels: Vec<u8>,
}

impl RasterImage {
    pub fn new(width: u32, height: u32, background: Rgba) -> Self {
        let mut pixels = vec![0u8; width as usize * height as usize * 4];
        for chunk in pixels.chunks_exact_mut(4) {
            chunk[0] = background.r;
            chunk[1] = background.g;
            chunk[2] = background.b;
            chunk[3] = background.a;
        }
        Self { width, height, pixels }
    }

    fn offset(&self, x: u32, y: u32) -> Option<usize> {
        if x >= self.width || y >= self.height {
            return None;
        }
        Some(((y * self.width + x) * 4) as usize)
    }

    pub fn blend_pixel(&mut self, x: i32, y: i32, color: Rgba) {
        if x < 0 || y < 0 {
            return;
        }
        let Some(offset) = self.offset(x as u32, y as u32) else {
            return;
        };
        if color.a == 255 {
            self.pixels[offset] = color.r;
            self.pixels[offset + 1] = color.g;
            self.pixels[offset + 2] = color.b;
            self.pixels[offset + 3] = 255;
            return;
        }
        let alpha = color.a as u16;
        let inv = 255u16.saturating_sub(alpha);
        self.pixels[offset] = (((color.r as u16 * alpha) + (self.pixels[offset] as u16 * inv)) / 255) as u8;
        self.pixels[offset + 1] = (((color.g as u16 * alpha) + (self.pixels[offset + 1] as u16 * inv)) / 255) as u8;
        self.pixels[offset + 2] = (((color.b as u16 * alpha) + (self.pixels[offset + 2] as u16 * inv)) / 255) as u8;
        self.pixels[offset + 3] = 255;
    }

    pub fn fill_rect(&mut self, x: i32, y: i32, width: i32, height: i32, color: Rgba) {
        if width <= 0 || height <= 0 {
            return;
        }
        let x0 = max(0, x);
        let y0 = max(0, y);
        let x1 = min(self.width as i32, x + width);
        let y1 = min(self.height as i32, y + height);
        for yy in y0..y1 {
            for xx in x0..x1 {
                self.blend_pixel(xx, yy, color);
            }
        }
    }

    pub fn stroke_rect(&mut self, x: i32, y: i32, width: i32, height: i32, thickness: i32, color: Rgba) {
        if width <= 0 || height <= 0 || thickness <= 0 {
            return;
        }
        self.fill_rect(x, y, width, thickness, color);
        self.fill_rect(x, y + height - thickness, width, thickness, color);
        self.fill_rect(x, y, thickness, height, color);
        self.fill_rect(x + width - thickness, y, thickness, height, color);
    }

    pub fn draw_line(&mut self, mut x0: i32, mut y0: i32, x1: i32, y1: i32, thickness: i32, color: Rgba) {
        let dx = (x1 - x0).abs();
        let sx = if x0 < x1 { 1 } else { -1 };
        let dy = -(y1 - y0).abs();
        let sy = if y0 < y1 { 1 } else { -1 };
        let mut err = dx + dy;
        let radius = max(1, thickness) / 2;
        loop {
            self.fill_rect(x0 - radius, y0 - radius, max(1, thickness), max(1, thickness), color);
            if x0 == x1 && y0 == y1 {
                break;
            }
            let e2 = err * 2;
            if e2 >= dy {
                err += dy;
                x0 += sx;
            }
            if e2 <= dx {
                err += dx;
                y0 += sy;
            }
        }
    }

    pub fn fill_ellipse(&mut self, cx: i32, cy: i32, rx: i32, ry: i32, color: Rgba) {
        if rx <= 0 || ry <= 0 {
            return;
        }
        for y in -ry..=ry {
            for x in -rx..=rx {
                let lhs = (x * x * ry * ry + y * y * rx * rx) as i64;
                let rhs = (rx * rx * ry * ry) as i64;
                if lhs <= rhs {
                    self.blend_pixel(cx + x, cy + y, color);
                }
            }
        }
    }

    pub fn stroke_ellipse(&mut self, cx: i32, cy: i32, rx: i32, ry: i32, thickness: i32, color: Rgba) {
        if rx <= 0 || ry <= 0 || thickness <= 0 {
            return;
        }
        let outer_rx = rx;
        let outer_ry = ry;
        let inner_rx = max(0, rx - thickness);
        let inner_ry = max(0, ry - thickness);
        for y in -outer_ry..=outer_ry {
            for x in -outer_rx..=outer_rx {
                let lhs = (x * x * outer_ry * outer_ry + y * y * outer_rx * outer_rx) as i64;
                let rhs = (outer_rx * outer_rx * outer_ry * outer_ry) as i64;
                if lhs > rhs {
                    continue;
                }
                let inner = (x * x * inner_ry * inner_ry + y * y * inner_rx * inner_rx) as i64;
                let inner_rhs = (inner_rx * inner_rx * inner_ry * inner_ry) as i64;
                if inner_rx > 0 && inner_ry > 0 && inner <= inner_rhs {
                    continue;
                }
                self.blend_pixel(cx + x, cy + y, color);
            }
        }
    }

    pub fn draw_text(&mut self, x: i32, y: i32, size_px: i32, color: Rgba, text: &str, bold: bool) {
        let scale = max(1, size_px / 8);
        let mut pen_x = x;
        let advance = 6 * scale;
        for ch in text.chars() {
            if ch == '\n' {
                pen_x = x;
                continue;
            }
            self.draw_glyph(pen_x, y, scale, color, ch);
            if bold {
                self.draw_glyph(pen_x + 1, y, scale, color, ch);
            }
            pen_x += advance;
        }
    }

    pub fn measure_text(&self, text: &str, size_px: i32) -> i32 {
        let scale = max(1, size_px / 8);
        text.chars().count() as i32 * 6 * scale
    }

    fn draw_glyph(&mut self, x: i32, y: i32, scale: i32, color: Rgba, ch: char) {
        let glyph = glyph_rows(ch);
        for (row, bits) in glyph.iter().enumerate() {
            for col in 0..5 {
                if bits & (1 << (4 - col)) != 0 {
                    self.fill_rect(x + col as i32 * scale, y + row as i32 * scale, scale, scale, color);
                }
            }
        }
    }

    pub fn encode_png(&self) -> Vec<u8> {
        let mut raw = Vec::with_capacity((self.width * self.height * 4 + self.height) as usize);
        for y in 0..self.height as usize {
            raw.push(0);
            let start = y * self.width as usize * 4;
            let end = start + self.width as usize * 4;
            raw.extend_from_slice(&self.pixels[start..end]);
        }
        let mut z = Vec::new();
        z.extend_from_slice(&[0x78, 0x01]);
        let mut remaining = raw.as_slice();
        while !remaining.is_empty() {
            let chunk_len = remaining.len().min(65_535);
            let final_block = chunk_len == remaining.len();
            z.push(if final_block { 0x01 } else { 0x00 });
            z.push((chunk_len & 0xFF) as u8);
            z.push(((chunk_len >> 8) & 0xFF) as u8);
            let nlen = !chunk_len as u16;
            z.push((nlen & 0xFF) as u8);
            z.push((nlen >> 8) as u8);
            z.extend_from_slice(&remaining[..chunk_len]);
            remaining = &remaining[chunk_len..];
        }
        let adler = adler32(&raw);
        z.extend_from_slice(&adler.to_be_bytes());

        let mut out = Vec::new();
        out.extend_from_slice(&[137, 80, 78, 71, 13, 10, 26, 10]);
        let mut ihdr = Vec::with_capacity(13);
        ihdr.extend_from_slice(&self.width.to_be_bytes());
        ihdr.extend_from_slice(&self.height.to_be_bytes());
        ihdr.extend_from_slice(&[8, 6, 0, 0, 0]);
        write_chunk(&mut out, b"IHDR", &ihdr);
        write_chunk(&mut out, b"IDAT", &z);
        write_chunk(&mut out, b"IEND", &[]);
        out
    }

    pub fn encode_jpeg(&self, quality: u8) -> Vec<u8> {
        let mut writer = JpegWriter::new(self.width as usize, self.height as usize, quality);
        writer.encode(&self.pixels)
    }
}

fn write_chunk(out: &mut Vec<u8>, tag: &[u8; 4], data: &[u8]) {
    out.extend_from_slice(&(data.len() as u32).to_be_bytes());
    out.extend_from_slice(tag);
    out.extend_from_slice(data);
    let mut crc_buf = Vec::with_capacity(tag.len() + data.len());
    crc_buf.extend_from_slice(tag);
    crc_buf.extend_from_slice(data);
    out.extend_from_slice(&crc32(&crc_buf).to_be_bytes());
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for &byte in bytes {
        crc ^= byte as u32;
        for _ in 0..8 {
            let mask = if crc & 1 == 1 { 0xEDB8_8320 } else { 0 };
            crc = (crc >> 1) ^ mask;
        }
    }
    !crc
}

fn adler32(bytes: &[u8]) -> u32 {
    const MOD: u32 = 65_521;
    let mut a = 1u32;
    let mut b = 0u32;
    for &byte in bytes {
        a = (a + byte as u32) % MOD;
        b = (b + a) % MOD;
    }
    (b << 16) | a
}

fn glyph_rows(ch: char) -> [u8; 7] {
    let c = if ch.is_ascii_lowercase() { ch.to_ascii_uppercase() } else { ch };
    match c {
        'A' => [0x0E, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
        'B' => [0x1E, 0x11, 0x11, 0x1E, 0x11, 0x11, 0x1E],
        'C' => [0x0E, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0E],
        'D' => [0x1E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1E],
        'E' => [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x1F],
        'F' => [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x10],
        'G' => [0x0E, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0E],
        'H' => [0x11, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
        'I' => [0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x1F],
        'J' => [0x07, 0x02, 0x02, 0x02, 0x12, 0x12, 0x0C],
        'K' => [0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11],
        'L' => [0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1F],
        'M' => [0x11, 0x1B, 0x15, 0x15, 0x11, 0x11, 0x11],
        'N' => [0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11],
        'O' => [0x0E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
        'P' => [0x1E, 0x11, 0x11, 0x1E, 0x10, 0x10, 0x10],
        'Q' => [0x0E, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0D],
        'R' => [0x1E, 0x11, 0x11, 0x1E, 0x14, 0x12, 0x11],
        'S' => [0x0F, 0x10, 0x10, 0x0E, 0x01, 0x01, 0x1E],
        'T' => [0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04],
        'U' => [0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
        'V' => [0x11, 0x11, 0x11, 0x11, 0x11, 0x0A, 0x04],
        'W' => [0x11, 0x11, 0x11, 0x15, 0x15, 0x15, 0x0A],
        'X' => [0x11, 0x11, 0x0A, 0x04, 0x0A, 0x11, 0x11],
        'Y' => [0x11, 0x11, 0x0A, 0x04, 0x04, 0x04, 0x04],
        'Z' => [0x1F, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1F],
        '0' => [0x0E, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0E],
        '1' => [0x04, 0x0C, 0x04, 0x04, 0x04, 0x04, 0x0E],
        '2' => [0x0E, 0x11, 0x01, 0x02, 0x04, 0x08, 0x1F],
        '3' => [0x1F, 0x02, 0x04, 0x02, 0x01, 0x11, 0x0E],
        '4' => [0x02, 0x06, 0x0A, 0x12, 0x1F, 0x02, 0x02],
        '5' => [0x1F, 0x10, 0x1E, 0x01, 0x01, 0x11, 0x0E],
        '6' => [0x06, 0x08, 0x10, 0x1E, 0x11, 0x11, 0x0E],
        '7' => [0x1F, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08],
        '8' => [0x0E, 0x11, 0x11, 0x0E, 0x11, 0x11, 0x0E],
        '9' => [0x0E, 0x11, 0x11, 0x0F, 0x01, 0x02, 0x0C],
        '!' => [0x04, 0x04, 0x04, 0x04, 0x04, 0x00, 0x04],
        '?' => [0x0E, 0x11, 0x01, 0x02, 0x04, 0x00, 0x04],
        '.' => [0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x06],
        ',' => [0x00, 0x00, 0x00, 0x00, 0x06, 0x06, 0x04],
        ':' => [0x00, 0x06, 0x06, 0x00, 0x06, 0x06, 0x00],
        ';' => [0x00, 0x06, 0x06, 0x00, 0x06, 0x06, 0x04],
        '-' => [0x00, 0x00, 0x00, 0x1F, 0x00, 0x00, 0x00],
        '_' => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x1F],
        '+' => [0x00, 0x04, 0x04, 0x1F, 0x04, 0x04, 0x00],
        '/' => [0x01, 0x02, 0x02, 0x04, 0x08, 0x08, 0x10],
        '\\' => [0x10, 0x08, 0x08, 0x04, 0x02, 0x02, 0x01],
        '(' => [0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02],
        ')' => [0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08],
        '[' => [0x0E, 0x08, 0x08, 0x08, 0x08, 0x08, 0x0E],
        ']' => [0x0E, 0x02, 0x02, 0x02, 0x02, 0x02, 0x0E],
        '&' => [0x0C, 0x12, 0x14, 0x08, 0x15, 0x12, 0x0D],
        '%' => [0x18, 0x19, 0x02, 0x04, 0x08, 0x13, 0x03],
        '*' => [0x00, 0x15, 0x0E, 0x1F, 0x0E, 0x15, 0x00],
        '=' => [0x00, 0x1F, 0x00, 0x1F, 0x00, 0x00, 0x00],
        '"' => [0x0A, 0x0A, 0x04, 0x00, 0x00, 0x00, 0x00],
        '\'' => [0x04, 0x04, 0x02, 0x00, 0x00, 0x00, 0x00],
        ' ' => [0x00; 7],
        _ => [0x1F, 0x11, 0x15, 0x15, 0x15, 0x11, 0x1F],
    }
}

const ZIGZAG: [usize; 64] = [
    0, 1, 5, 6, 14, 15, 27, 28,
    2, 4, 7, 13, 16, 26, 29, 42,
    3, 8, 12, 17, 25, 30, 41, 43,
    9, 11, 18, 24, 31, 40, 44, 53,
    10, 19, 23, 32, 39, 45, 52, 54,
    20, 22, 33, 38, 46, 51, 55, 60,
    21, 34, 37, 47, 50, 56, 59, 61,
    35, 36, 48, 49, 57, 58, 62, 63,
];

const STD_LUMA_Q: [u8; 64] = [
    16, 11, 10, 16, 24, 40, 51, 61,
    12, 12, 14, 19, 26, 58, 60, 55,
    14, 13, 16, 24, 40, 57, 69, 56,
    14, 17, 22, 29, 51, 87, 80, 62,
    18, 22, 37, 56, 68, 109, 103, 77,
    24, 35, 55, 64, 81, 104, 113, 92,
    49, 64, 78, 87, 103, 121, 120, 101,
    72, 92, 95, 98, 112, 100, 103, 99,
];

const STD_CHROMA_Q: [u8; 64] = [
    17, 18, 24, 47, 99, 99, 99, 99,
    18, 21, 26, 66, 99, 99, 99, 99,
    24, 26, 56, 99, 99, 99, 99, 99,
    47, 66, 99, 99, 99, 99, 99, 99,
    99, 99, 99, 99, 99, 99, 99, 99,
    99, 99, 99, 99, 99, 99, 99, 99,
    99, 99, 99, 99, 99, 99, 99, 99,
    99, 99, 99, 99, 99, 99, 99, 99,
];

const BITS_DC_LUMA: [u8; 17] = [0, 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0];
const VAL_DC_LUMA: [u8; 12] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
const BITS_AC_LUMA: [u8; 17] = [0, 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d];
const VAL_AC_LUMA: [u8; 162] = [
    0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61,
    0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08, 0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52,
    0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25,
    0x26, 0x27, 0x28, 0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45,
    0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64,
    0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x83,
    0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99,
    0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6,
    0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3,
    0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8,
    0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA,
];
const BITS_DC_CHROMA: [u8; 17] = [0, 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0];
const VAL_DC_CHROMA: [u8; 12] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
const BITS_AC_CHROMA: [u8; 17] = [0, 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 0x77];
const VAL_AC_CHROMA: [u8; 162] = [
    0x00, 0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61,
    0x71, 0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xA1, 0xB1, 0xC1, 0x09, 0x23, 0x33,
    0x52, 0xF0, 0x15, 0x62, 0x72, 0xD1, 0x0A, 0x16, 0x24, 0x34, 0xE1, 0x25, 0xF1, 0x17, 0x18,
    0x19, 0x1A, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44,
    0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63,
    0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A,
    0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97,
    0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4,
    0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA,
    0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7,
    0xE8, 0xE9, 0xEA, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA,
];

#[derive(Clone)]
struct HuffmanTable {
    codes: [u16; 256],
    sizes: [u8; 256],
}

impl HuffmanTable {
    fn from_spec(bits: &[u8; 17], values: &[u8]) -> Self {
        let mut codes = [0u16; 256];
        let mut sizes = [0u8; 256];
        let mut code = 0u16;
        let mut k = 0usize;
        for bit_len in 1..=16 {
            for _ in 0..bits[bit_len] {
                let value = values[k] as usize;
                codes[value] = code;
                sizes[value] = bit_len as u8;
                code += 1;
                k += 1;
            }
            code <<= 1;
        }
        Self { codes, sizes }
    }
}

struct BitWriter {
    bytes: Vec<u8>,
    cur: u32,
    bits: u8,
}

impl BitWriter {
    fn new() -> Self {
        Self { bytes: Vec::new(), cur: 0, bits: 0 }
    }

    fn write(&mut self, code: u16, size: u8) {
        if size == 0 { return; }
        self.cur |= (code as u32) << (24 - self.bits as u32 - size as u32);
        self.bits += size;
        while self.bits >= 8 {
            let byte = ((self.cur >> 16) & 0xFF) as u8;
            self.bytes.push(byte);
            if byte == 0xFF {
                self.bytes.push(0x00);
            }
            self.cur <<= 8;
            self.bits -= 8;
        }
    }

    fn flush(&mut self) {
        if self.bits > 0 {
            let byte = ((self.cur >> 16) & 0xFF) as u8;
            self.bytes.push(byte);
            if byte == 0xFF {
                self.bytes.push(0x00);
            }
            self.cur = 0;
            self.bits = 0;
        }
    }
}

struct JpegWriter {
    width: usize,
    height: usize,
    quality: u8,
}

impl JpegWriter {
    fn new(width: usize, height: usize, quality: u8) -> Self {
        Self { width, height, quality: quality.max(1).min(100) }
    }

    fn encode(&mut self, rgba: &[u8]) -> Vec<u8> {
        let dc_luma = HuffmanTable::from_spec(&BITS_DC_LUMA, &VAL_DC_LUMA);
        let ac_luma = HuffmanTable::from_spec(&BITS_AC_LUMA, &VAL_AC_LUMA);
        let dc_chroma = HuffmanTable::from_spec(&BITS_DC_CHROMA, &VAL_DC_CHROMA);
        let ac_chroma = HuffmanTable::from_spec(&BITS_AC_CHROMA, &VAL_AC_CHROMA);
        let qy = scaled_qtable(&STD_LUMA_Q, self.quality);
        let qc = scaled_qtable(&STD_CHROMA_Q, self.quality);

        let mut out = Vec::new();
        out.extend_from_slice(&[0xFF, 0xD8]);
        write_app0(&mut out);
        write_dqt(&mut out, 0, &qy);
        write_dqt(&mut out, 1, &qc);
        write_sof0(&mut out, self.width as u16, self.height as u16);
        write_dht(&mut out, 0x00, &BITS_DC_LUMA, &VAL_DC_LUMA);
        write_dht(&mut out, 0x10, &BITS_AC_LUMA, &VAL_AC_LUMA);
        write_dht(&mut out, 0x01, &BITS_DC_CHROMA, &VAL_DC_CHROMA);
        write_dht(&mut out, 0x11, &BITS_AC_CHROMA, &VAL_AC_CHROMA);
        write_sos(&mut out);

        let mut bits = BitWriter::new();
        let mut prev_y = 0i32;
        let mut prev_cb = 0i32;
        let mut prev_cr = 0i32;
        let padded_w = (self.width + 7) & !7;
        let padded_h = (self.height + 7) & !7;
        for by in (0..padded_h).step_by(8) {
            for bx in (0..padded_w).step_by(8) {
                let y = sample_block(rgba, self.width, self.height, bx, by, Component::Y);
                let cb = sample_block(rgba, self.width, self.height, bx, by, Component::Cb);
                let cr = sample_block(rgba, self.width, self.height, bx, by, Component::Cr);
                prev_y = write_block(&mut bits, &y, &qy, prev_y, &dc_luma, &ac_luma);
                prev_cb = write_block(&mut bits, &cb, &qc, prev_cb, &dc_chroma, &ac_chroma);
                prev_cr = write_block(&mut bits, &cr, &qc, prev_cr, &dc_chroma, &ac_chroma);
            }
        }
        bits.flush();
        out.extend_from_slice(&bits.bytes);
        out.extend_from_slice(&[0xFF, 0xD9]);
        out
    }
}

#[derive(Clone, Copy)]
enum Component { Y, Cb, Cr }

fn scaled_qtable(base: &[u8; 64], quality: u8) -> [u8; 64] {
    let q = quality.max(1).min(100);
    let scale = if q < 50 { 5000 / q as i32 } else { 200 - q as i32 * 2 };
    let mut out = [0u8; 64];
    for (index, value) in base.iter().enumerate() {
        let scaled = ((*value as i32 * scale + 50) / 100).clamp(1, 255);
        out[index] = scaled as u8;
    }
    out
}

fn write_marker_segment(out: &mut Vec<u8>, marker: u8, payload: &[u8]) {
    out.extend_from_slice(&[0xFF, marker]);
    out.extend_from_slice(&((payload.len() + 2) as u16).to_be_bytes());
    out.extend_from_slice(payload);
}

fn write_app0(out: &mut Vec<u8>) {
    write_marker_segment(out, 0xE0, &[0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00]);
}

fn write_dqt(out: &mut Vec<u8>, id: u8, table: &[u8; 64]) {
    let mut payload = Vec::with_capacity(65);
    payload.push(id);
    for &i in &ZIGZAG {
        payload.push(table[i]);
    }
    write_marker_segment(out, 0xDB, &payload);
}

fn write_sof0(out: &mut Vec<u8>, width: u16, height: u16) {
    let payload = [
        8,
        (height >> 8) as u8, height as u8,
        (width >> 8) as u8, width as u8,
        3,
        1, 0x11, 0,
        2, 0x11, 1,
        3, 0x11, 1,
    ];
    write_marker_segment(out, 0xC0, &payload);
}

fn write_dht(out: &mut Vec<u8>, id: u8, bits: &[u8; 17], values: &[u8]) {
    let mut payload = Vec::with_capacity(1 + 16 + values.len());
    payload.push(id);
    payload.extend_from_slice(&bits[1..]);
    payload.extend_from_slice(values);
    write_marker_segment(out, 0xC4, &payload);
}

fn write_sos(out: &mut Vec<u8>) {
    let payload = [3, 1, 0x00, 2, 0x11, 3, 0x11, 0, 63, 0];
    write_marker_segment(out, 0xDA, &payload);
}

fn sample_block(rgba: &[u8], width: usize, height: usize, bx: usize, by: usize, component: Component) -> [f32; 64] {
    let mut out = [0f32; 64];
    for y in 0..8 {
        for x in 0..8 {
            let px = min(width.saturating_sub(1), bx + x);
            let py = min(height.saturating_sub(1), by + y);
            let offset = (py * width + px) * 4;
            let r = rgba[offset] as f32;
            let g = rgba[offset + 1] as f32;
            let b = rgba[offset + 2] as f32;
            let value = match component {
                Component::Y => 0.299 * r + 0.587 * g + 0.114 * b - 128.0,
                Component::Cb => -0.168736 * r - 0.331264 * g + 0.5 * b,
                Component::Cr => 0.5 * r - 0.418688 * g - 0.081312 * b,
            };
            out[y * 8 + x] = value;
        }
    }
    out
}

fn fdct(input: &[f32; 64]) -> [f32; 64] {
    let mut out = [0f32; 64];
    for v in 0..8 {
        for u in 0..8 {
            let mut sum = 0f32;
            for y in 0..8 {
                for x in 0..8 {
                    let s = input[y * 8 + x];
                    let cu = ((2 * x + 1) as f32 * u as f32 * std::f32::consts::PI / 16.0).cos();
                    let cv = ((2 * y + 1) as f32 * v as f32 * std::f32::consts::PI / 16.0).cos();
                    sum += s * cu * cv;
                }
            }
            let au = if u == 0 { 1.0 / 2f32.sqrt() } else { 1.0 };
            let av = if v == 0 { 1.0 / 2f32.sqrt() } else { 1.0 };
            out[v * 8 + u] = 0.25 * au * av * sum;
        }
    }
    out
}

fn write_block(bits: &mut BitWriter, block: &[f32; 64], qtable: &[u8; 64], prev_dc: i32, dc_table: &HuffmanTable, ac_table: &HuffmanTable) -> i32 {
    let dct = fdct(block);
    let mut quant = [0i32; 64];
    for i in 0..64 {
        quant[i] = (dct[i] / qtable[i] as f32).round() as i32;
    }
    let dc = quant[0];
    let diff = dc - prev_dc;
    let dc_cat = magnitude_category(diff);
    write_huff(bits, dc_table, dc_cat as u8);
    if dc_cat > 0 {
        let extra = magnitude_bits(diff, dc_cat);
        bits.write(extra as u16, dc_cat as u8);
    }

    let mut run = 0u8;
    for idx in 1..64 {
        let value = quant[ZIGZAG[idx]];
        if value == 0 {
            run = run.saturating_add(1);
            continue;
        }
        while run >= 16 {
            write_huff(bits, ac_table, 0xF0);
            run -= 16;
        }
        let cat = magnitude_category(value) as u8;
        write_huff(bits, ac_table, (run << 4) | cat);
        let extra = magnitude_bits(value, cat as i32);
        bits.write(extra as u16, cat);
        run = 0;
    }
    if run > 0 {
        write_huff(bits, ac_table, 0x00);
    }
    dc
}

fn write_huff(bits: &mut BitWriter, table: &HuffmanTable, symbol: u8) {
    bits.write(table.codes[symbol as usize], table.sizes[symbol as usize]);
}

fn magnitude_category(value: i32) -> i32 {
    if value == 0 {
        return 0;
    }
    32 - value.abs().leading_zeros() as i32
}

fn magnitude_bits(value: i32, size: i32) -> u32 {
    if size <= 0 { return 0; }
    if value >= 0 {
        value as u32
    } else {
        ((1 << size) - 1 + value) as u32
    }
}
