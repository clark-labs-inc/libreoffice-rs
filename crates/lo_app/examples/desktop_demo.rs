//! Tiny lo_app demo: open a Writer template, type into it, save it, and
//! render an HTML "window shell" + a "start center" landing page.
//!
//! Run with: `cargo run -p lo_app --example desktop_demo`

use std::collections::BTreeMap;

use lo_app::DesktopApp;
use lo_uno::UnoValue;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let app = DesktopApp::new("target/demo-profile")?;
    let writer = app.open_template("writer:report")?;
    app.execute_window_command(
        writer,
        ".uno:InsertText",
        &BTreeMap::from([(
            "text".to_string(),
            UnoValue::String("Hello from the desktop shell".to_string()),
        )]),
    )?;
    app.save_window(writer)?;
    app.save_window_shell(writer, "target/demo-profile/shells/writer.html")?;
    app.save_start_center("target/demo-profile/start-center.html")?;
    println!("wrote target/demo-profile/{{start-center.html, shells/writer.html, exports/...}}");
    Ok(())
}
