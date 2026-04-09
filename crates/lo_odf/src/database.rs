use std::path::Path;

use lo_core::{ColumnType, DatabaseDocument, DbValue, Result};

use crate::common::{package_database_document, ExtraFile};

fn sdbc_type(column_type: &ColumnType) -> &'static str {
    match column_type {
        ColumnType::Integer => "integer",
        ColumnType::Float => "double",
        ColumnType::Text => "varchar",
        ColumnType::Bool => "boolean",
    }
}

pub fn serialize_database_document(database: &DatabaseDocument) -> String {
    // Mirror the shape that LibreOffice itself writes for "connect to existing
    // CSV" Base documents: an office:database with a flat-file data-source
    // pointing at the ./database/ subdirectory that we also pack into the
    // archive. This structure was derived from the evolocal.odb / biblio.odb
    // presets shipped with LibreOffice 26.
    let mut out = String::new();
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    out.push_str("<office:document-content");
    out.push_str(" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"");
    out.push_str(" xmlns:db=\"urn:oasis:names:tc:opendocument:xmlns:database:1.0\"");
    out.push_str(" xmlns:xlink=\"http://www.w3.org/1999/xlink\"");
    out.push_str(" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"");
    out.push_str(" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"");
    out.push_str(" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"");
    out.push_str(" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\"");
    out.push_str(" office:version=\"1.3\">\n");
    out.push_str("  <office:automatic-styles/>\n");
    out.push_str("  <office:body>\n");
    out.push_str("    <office:database>\n");
    out.push_str("      <db:data-source>\n");
    out.push_str("        <db:connection-data>\n");
    out.push_str("          <db:database-description>\n");
    out.push_str("            <db:file-based-database xlink:href=\"./database\" db:media-type=\"text/csv\" db:extension=\"csv\"/>\n");
    out.push_str("          </db:database-description>\n");
    out.push_str("          <db:login db:is-password-required=\"false\"/>\n");
    out.push_str("        </db:connection-data>\n");
    out.push_str("        <db:driver-settings db:system-driver-settings=\"\" db:base-dn=\"\" db:parameter-name-substitution=\"false\">\n");
    out.push_str("          <db:font-charset db:encoding=\"utf-8\"/>\n");
    out.push_str("          <db:delimiter db:field=\",\" db:string=\"&quot;\" db:decimal=\".\" db:thousand=\",\"/>\n");
    out.push_str("        </db:driver-settings>\n");
    out.push_str("        <db:application-connection-settings db:is-table-name-length-limited=\"false\" db:append-table-alias-name=\"false\">\n");
    out.push_str("          <db:table-filter>\n");
    out.push_str("            <db:table-include-filter>\n");
    out.push_str("              <db:table-filter-pattern>%</db:table-filter-pattern>\n");
    out.push_str("            </db:table-include-filter>\n");
    out.push_str("          </db:table-filter>\n");
    out.push_str("          <db:data-source-settings>\n");
    out.push_str("            <db:data-source-setting db:data-source-setting-is-list=\"false\" db:data-source-setting-name=\"HeaderLine\" db:data-source-setting-type=\"boolean\"><db:data-source-setting-value>true</db:data-source-setting-value></db:data-source-setting>\n");
    out.push_str("            <db:data-source-setting db:data-source-setting-is-list=\"false\" db:data-source-setting-name=\"CharSet\" db:data-source-setting-type=\"string\"><db:data-source-setting-value>utf-8</db:data-source-setting-value></db:data-source-setting>\n");
    out.push_str("            <db:data-source-setting db:data-source-setting-is-list=\"false\" db:data-source-setting-name=\"Extension\" db:data-source-setting-type=\"string\"><db:data-source-setting-value>csv</db:data-source-setting-value></db:data-source-setting>\n");
    out.push_str("            <db:data-source-setting db:data-source-setting-is-list=\"false\" db:data-source-setting-name=\"FieldDelimiter\" db:data-source-setting-type=\"string\"><db:data-source-setting-value>,</db:data-source-setting-value></db:data-source-setting>\n");
    out.push_str("          </db:data-source-settings>\n");
    out.push_str("        </db:application-connection-settings>\n");
    out.push_str("      </db:data-source>\n");
    out.push_str("      <db:table-representations>\n");
    for table in &database.tables {
        out.push_str(&format!(
            "        <db:table-representation db:name=\"{}\"/>\n",
            lo_core::escape_attr(&table.name)
        ));
    }
    out.push_str("      </db:table-representations>\n");
    out.push_str("    </office:database>\n");
    out.push_str("  </office:body>\n");
    out.push_str("</office:document-content>\n");
    let _ = (DbValue::Null, ColumnType::Text, sdbc_type); // suppress unused warnings
    out
}

fn csv_for_table(table: &lo_core::TableData) -> Vec<u8> {
    let mut out = String::new();
    out.push_str(
        &table
            .columns
            .iter()
            .map(|column| column.name.replace('"', "\"\""))
            .collect::<Vec<_>>()
            .join(","),
    );
    out.push('\n');
    for row in &table.rows {
        let cols: Vec<String> = row
            .iter()
            .map(|value| match value {
                DbValue::Null => String::new(),
                DbValue::Integer(v) => v.to_string(),
                DbValue::Float(v) => v.to_string(),
                DbValue::Text(v) => {
                    if v.contains(',') || v.contains('"') || v.contains('\n') {
                        format!("\"{}\"", v.replace('"', "\"\""))
                    } else {
                        v.clone()
                    }
                }
                DbValue::Bool(v) => v.to_string(),
            })
            .collect();
        out.push_str(&cols.join(","));
        out.push('\n');
    }
    out.into_bytes()
}

pub fn save_database_document(path: impl AsRef<Path>, database: &DatabaseDocument) -> Result<()> {
    let content = serialize_database_document(database);
    let mut extras = Vec::new();
    for table in &database.tables {
        extras.push(ExtraFile::new(
            format!("database/{}.csv", table.name),
            "text/csv",
            csv_for_table(table),
        ));
    }
    package_database_document(path, content, extras)
}
