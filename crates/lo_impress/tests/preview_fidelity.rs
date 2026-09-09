use lo_core::{Presentation, RasterImage, Rgba, Slide};

#[test]
fn jpeg_preserves_directional_detail() {
    let mut source = RasterImage::new(65, 49, Rgba::WHITE);
    for y in 0..source.height {
        for x in 0..source.width {
            let value = ((x * 3 + y) % 256) as u8;
            source.blend_pixel(x as i32, y as i32, Rgba::rgba(value, value, value, 255));
        }
    }
    let decoded = image::load_from_memory(&source.encode_jpeg(95)).unwrap().to_rgb8();
    assert_eq!(decoded.dimensions(), (65, 49));
    let error: u64 = decoded.pixels().zip(source.pixels.chunks_exact(4))
        .map(|(actual, expected)| actual.0.iter().zip(expected).take(3)
            .map(|(a, b)| (*a as i32 - *b as i32).unsigned_abs() as u64).sum::<u64>())
        .sum();
    let mean_error = error as f64 / (65.0 * 49.0 * 3.0);
    assert!(mean_error < 2.0, "JPEG mean channel error: {mean_error}");
}

#[test]
fn slide_metadata_does_not_paint_over_slide_content() {
    let mut deck = Presentation::new("Deck");
    deck.slides.push(Slide::default());
    let raster = lo_impress::render_png_pages(&deck, 96);
    let pdf = lo_impress::to_pdf(&deck);
    deck.slides[0].name = "Internal slide name".into();
    deck.slides[0].notes = vec!["Private speaker notes".into()];
    assert!(lo_impress::render_png_pages(&deck, 96) == raster, "metadata changed slide pixels");
    assert!(lo_impress::to_pdf(&deck) == pdf, "metadata changed PDF slide content");
}
