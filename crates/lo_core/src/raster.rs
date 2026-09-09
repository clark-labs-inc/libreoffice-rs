use std::cmp::{max, min};

use crate::{LoError, Result};

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
            let r = chars
                .next()
                .and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
            let g = chars
                .next()
                .and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
            let b = chars
                .next()
                .and_then(|ch| u8::from_str_radix(&format!("{ch}{ch}"), 16).ok());
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

/// Decode a WebP image using the pure-Rust `image-webp` codec exposed by
/// the `image` crate. Animated WebP files resolve to their first frame,
/// matching the static-image behavior used by PDF readers.
pub fn decode_webp(bytes: &[u8]) -> Result<RasterImage> {
    let decoded = image::load_from_memory_with_format(bytes, image::ImageFormat::WebP)
        .map_err(|error| LoError::Parse(format!("invalid WebP image: {error}")))?
        .to_rgba8();
    let (width, height) = decoded.dimensions();
    if width == 0 || height == 0 {
        return Err(LoError::InvalidInput(
            "WebP dimensions must be non-zero".to_string(),
        ));
    }
    Ok(RasterImage {
        width,
        height,
        pixels: decoded.into_raw(),
    })
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
        Self {
            width,
            height,
            pixels,
        }
    }

    pub fn encode_webp(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        image::codecs::webp::WebPEncoder::new_lossless(&mut bytes)
            .encode(
                &self.pixels,
                self.width,
                self.height,
                image::ExtendedColorType::Rgba8,
            )
            .map_err(|error| LoError::Parse(format!("could not encode WebP image: {error}")))?;
        Ok(bytes)
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
        self.pixels[offset] =
            (((color.r as u16 * alpha) + (self.pixels[offset] as u16 * inv)) / 255) as u8;
        self.pixels[offset + 1] =
            (((color.g as u16 * alpha) + (self.pixels[offset + 1] as u16 * inv)) / 255) as u8;
        self.pixels[offset + 2] =
            (((color.b as u16 * alpha) + (self.pixels[offset + 2] as u16 * inv)) / 255) as u8;
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

    pub fn stroke_rect(
        &mut self,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        thickness: i32,
        color: Rgba,
    ) {
        if width <= 0 || height <= 0 || thickness <= 0 {
            return;
        }
        self.fill_rect(x, y, width, thickness, color);
        self.fill_rect(x, y + height - thickness, width, thickness, color);
        self.fill_rect(x, y, thickness, height, color);
        self.fill_rect(x + width - thickness, y, thickness, height, color);
    }

    pub fn draw_line(
        &mut self,
        mut x0: i32,
        mut y0: i32,
        x1: i32,
        y1: i32,
        thickness: i32,
        color: Rgba,
    ) {
        let dx = (x1 - x0).abs();
        let sx = if x0 < x1 { 1 } else { -1 };
        let dy = -(y1 - y0).abs();
        let sy = if y0 < y1 { 1 } else { -1 };
        let mut err = dx + dy;
        let radius = max(1, thickness) / 2;
        loop {
            self.fill_rect(
                x0 - radius,
                y0 - radius,
                max(1, thickness),
                max(1, thickness),
                color,
            );
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

    pub fn stroke_ellipse(
        &mut self,
        cx: i32,
        cy: i32,
        rx: i32,
        ry: i32,
        thickness: i32,
        color: Rgba,
    ) {
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
                    self.fill_rect(
                        x + col as i32 * scale,
                        y + row as i32 * scale,
                        scale,
                        scale,
                        color,
                    );
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
        let rgb: Vec<u8> = self.pixels.chunks_exact(4)
            .flat_map(|pixel| pixel[..3].iter().copied()).collect();
        let mut bytes = Vec::new();
        image::codecs::jpeg::JpegEncoder::new_with_quality(&mut bytes, quality.clamp(1, 100))
            .encode(&rgb, self.width, self.height, image::ExtendedColorType::Rgb8)
            .expect("valid raster dimensions and RGB pixels");
        bytes
    }
}

/// Decode a non-interlaced, 8-bit PNG into RGBA pixels. Supports grayscale,
/// RGB, indexed color, grayscale+alpha, and RGBA, including every standard
/// PNG scanline filter.
pub fn decode_png(bytes: &[u8]) -> Result<RasterImage> {
    const SIGNATURE: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if bytes.get(..8) != Some(SIGNATURE) {
        return Err(LoError::Parse("invalid PNG signature".to_string()));
    }
    let mut offset = 8usize;
    let mut width = 0u32;
    let mut height = 0u32;
    let mut bit_depth = 0u8;
    let mut color_type = 0u8;
    let mut interlace = 0u8;
    let mut palette = Vec::new();
    let mut transparency = Vec::new();
    let mut compressed = Vec::new();
    while offset + 12 <= bytes.len() {
        let len = u32::from_be_bytes(bytes[offset..offset + 4].try_into().unwrap()) as usize;
        let tag = &bytes[offset + 4..offset + 8];
        let data_start = offset + 8;
        let data_end = data_start
            .checked_add(len)
            .ok_or_else(|| LoError::Parse("PNG chunk length overflow".to_string()))?;
        if data_end + 4 > bytes.len() {
            return Err(LoError::Parse("truncated PNG chunk".to_string()));
        }
        let data = &bytes[data_start..data_end];
        match tag {
            b"IHDR" if data.len() == 13 => {
                width = u32::from_be_bytes(data[0..4].try_into().unwrap());
                height = u32::from_be_bytes(data[4..8].try_into().unwrap());
                bit_depth = data[8];
                color_type = data[9];
                interlace = data[12];
            }
            b"PLTE" => palette.extend_from_slice(data),
            b"tRNS" => transparency.extend_from_slice(data),
            b"IDAT" => compressed.extend_from_slice(data),
            b"IEND" => break,
            _ => {}
        }
        offset = data_end + 4;
    }
    if width == 0 || height == 0 || bit_depth != 8 || interlace != 0 {
        return Err(LoError::Unsupported(
            "PNG images must be non-interlaced and 8-bit".to_string(),
        ));
    }
    let channels = match color_type {
        0 | 3 => 1usize,
        2 => 3,
        4 => 2,
        6 => 4,
        _ => {
            return Err(LoError::Unsupported(format!(
                "PNG color type {color_type} is not supported"
            )))
        }
    };
    let stride = (width as usize)
        .checked_mul(channels)
        .ok_or_else(|| LoError::InvalidInput("PNG dimensions are too large".to_string()))?;
    let expected = (stride + 1)
        .checked_mul(height as usize)
        .ok_or_else(|| LoError::InvalidInput("PNG dimensions are too large".to_string()))?;
    if expected > 512 * 1024 * 1024 {
        return Err(LoError::InvalidInput("PNG image is too large".to_string()));
    }
    let raw = crate::pdf::decode_flate_stream(&compressed)?;
    if raw.len() < expected {
        return Err(LoError::Parse("truncated PNG image data".to_string()));
    }
    let mut scanlines = vec![0u8; stride * height as usize];
    for row in 0..height as usize {
        let source = &raw[row * (stride + 1) + 1..row * (stride + 1) + 1 + stride];
        let filter = raw[row * (stride + 1)];
        let (before, current_and_after) = scanlines.split_at_mut(row * stride);
        let current = &mut current_and_after[..stride];
        let previous = if row == 0 {
            None
        } else {
            Some(&before[(row - 1) * stride..row * stride])
        };
        for index in 0..stride {
            let left = if index >= channels {
                current[index - channels]
            } else {
                0
            };
            let up = previous.map(|line| line[index]).unwrap_or(0);
            let upper_left = if index >= channels {
                previous.map(|line| line[index - channels]).unwrap_or(0)
            } else {
                0
            };
            current[index] = source[index].wrapping_add(match filter {
                0 => 0,
                1 => left,
                2 => up,
                3 => ((left as u16 + up as u16) / 2) as u8,
                4 => paeth(left, up, upper_left),
                other => return Err(LoError::Parse(format!("invalid PNG filter {other}"))),
            });
        }
    }
    let mut pixels = Vec::with_capacity(width as usize * height as usize * 4);
    for pixel in scanlines.chunks_exact(channels) {
        match color_type {
            0 => pixels.extend_from_slice(&[pixel[0], pixel[0], pixel[0], 255]),
            2 => pixels.extend_from_slice(&[pixel[0], pixel[1], pixel[2], 255]),
            3 => {
                let index = pixel[0] as usize;
                let start = index * 3;
                let rgb = palette.get(start..start + 3).ok_or_else(|| {
                    LoError::Parse("PNG palette index is out of range".to_string())
                })?;
                pixels.extend_from_slice(&[
                    rgb[0],
                    rgb[1],
                    rgb[2],
                    transparency.get(index).copied().unwrap_or(255),
                ]);
            }
            4 => pixels.extend_from_slice(&[pixel[0], pixel[0], pixel[0], pixel[1]]),
            6 => pixels.extend_from_slice(pixel),
            _ => unreachable!(),
        }
    }
    Ok(RasterImage {
        width,
        height,
        pixels,
    })
}

fn paeth(left: u8, up: u8, upper_left: u8) -> u8 {
    let p = left as i32 + up as i32 - upper_left as i32;
    let pa = (p - left as i32).abs();
    let pb = (p - up as i32).abs();
    let pc = (p - upper_left as i32).abs();
    if pa <= pb && pa <= pc {
        left
    } else if pb <= pc {
        up
    } else {
        upper_left
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
    let c = if ch.is_ascii_lowercase() {
        ch.to_ascii_uppercase()
    } else {
        ch
    };
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
