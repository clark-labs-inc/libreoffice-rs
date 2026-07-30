use lo_core::{LoError, Result};

const MIN_RASTER_DPI: u32 = 72;
const MAX_RASTER_DPI: u32 = 300;

fn raster_dpi(dpi: u32) -> Result<u32> {
    if dpi > MAX_RASTER_DPI {
        return Err(LoError::InvalidInput(format!(
            "raster DPI {dpi} exceeds the {MAX_RASTER_DPI} DPI memory-safety limit"
        )));
    }
    Ok(dpi.max(MIN_RASTER_DPI))
}

/// Rasterize a DOCX document directly to PNG pages at the requested DPI.
pub fn docx_to_png_pages(input: &[u8], dpi: u32) -> Result<Vec<Vec<u8>>> {
    let dpi = raster_dpi(dpi)?;
    let doc = lo_writer::from_docx_bytes("document", input)?;
    Ok(lo_writer::render_png_pages(&doc, dpi))
}

/// Rasterize a DOCX document directly to JPEG pages at the requested DPI.
pub fn docx_to_jpeg_pages(input: &[u8], dpi: u32, quality: u8) -> Result<Vec<Vec<u8>>> {
    let dpi = raster_dpi(dpi)?;
    let doc = lo_writer::from_docx_bytes("document", input)?;
    Ok(lo_writer::render_jpeg_pages(&doc, dpi, quality.max(1)))
}

/// Rasterize a PPTX deck directly to PNG slide images at the requested DPI.
pub fn pptx_to_png_pages(input: &[u8], dpi: u32) -> Result<Vec<Vec<u8>>> {
    let dpi = raster_dpi(dpi)?;
    let deck = lo_impress::from_pptx_bytes("presentation", input)?;
    Ok(lo_impress::render_png_pages(&deck, dpi))
}

/// Rasterize a PPTX deck directly to JPEG slide images at the requested DPI.
pub fn pptx_to_jpeg_pages(input: &[u8], dpi: u32, quality: u8) -> Result<Vec<Vec<u8>>> {
    let dpi = raster_dpi(dpi)?;
    let deck = lo_impress::from_pptx_bytes("presentation", input)?;
    Ok(lo_impress::render_jpeg_pages(&deck, dpi, quality.max(1)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizes_low_dpi_to_supported_minimum() {
        assert_eq!(raster_dpi(0).unwrap(), MIN_RASTER_DPI);
        assert_eq!(raster_dpi(71).unwrap(), MIN_RASTER_DPI);
        assert_eq!(raster_dpi(150).unwrap(), 150);
    }

    #[test]
    fn rejects_dpi_that_can_exhaust_page_memory_before_parsing() {
        let error = docx_to_png_pages(b"not a docx", MAX_RASTER_DPI + 1).unwrap_err();
        assert!(
            error.to_string().contains("memory-safety limit"),
            "unexpected error: {error}"
        );
    }
}
