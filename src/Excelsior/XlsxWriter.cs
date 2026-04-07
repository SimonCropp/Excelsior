using System.IO.Compression;
using System.Xml;

namespace Excelsior;

static class XlsxWriter
{
    const string NsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string NsPackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    const string NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    static readonly XmlWriterSettings xmlSettings = new()
    {
        Encoding = Encoding.UTF8
    };

    internal static void Write(ZipArchive archive, List<SheetData> sheets, StyleManager styleManager)
    {
        WriteEntry(archive, "[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", NsContentTypes);
            WriteDefault(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            WriteDefault(writer, "xml", "application/xml");
            WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            for (var i = 0; i < sheets.Count; i++)
            {
                WriteOverride(writer, $"/xl/worksheets/sheet{i + 1}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            }

            writer.WriteEndElement();
        });

        WriteEntry(archive, "_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", NsPackageRelationships);
            WriteRelationship(writer, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
            writer.WriteEndElement();
        });

        WriteEntry(archive, "xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", NsSpreadsheet);
            writer.WriteAttributeString("xmlns", "r", null, NsRelationships);
            writer.WriteStartElement("sheets");
            for (var i = 0; i < sheets.Count; i++)
            {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", sheets[i].Name);
                writer.WriteAttributeString("sheetId", (i + 1).ToString());
                writer.WriteAttributeString("r", "id", NsRelationships, $"rId{i + 1}");
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        });

        WriteEntry(archive, "xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", NsPackageRelationships);
            for (var i = 0; i < sheets.Count; i++)
            {
                WriteRelationship(writer, $"rId{i + 1}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $"worksheets/sheet{i + 1}.xml");
            }

            WriteRelationship(writer, $"rId{sheets.Count + 1}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
            writer.WriteEndElement();
        });

        WriteEntry(archive, "xl/styles.xml", styleManager.WriteStylesXml);

        for (var i = 0; i < sheets.Count; i++)
        {
            WriteSheet(archive, sheets[i], i + 1);
            if (sheets[i].Hyperlinks.Count > 0)
            {
                WriteSheetRels(archive, sheets[i], i + 1);
            }
        }
    }

    static void WriteSheet(ZipArchive archive, SheetData sheet, int sheetIndex) =>
        WriteEntry(archive, $"xl/worksheets/sheet{sheetIndex}.xml", writer =>
        {
            writer.WriteStartElement("worksheet", NsSpreadsheet);
            writer.WriteAttributeString("xmlns", "r", null, NsRelationships);

            // Frozen header pane
            writer.WriteStartElement("sheetViews");
            writer.WriteStartElement("sheetView");
            writer.WriteAttributeString("tabSelected", "1");
            writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteStartElement("pane");
            writer.WriteAttributeString("ySplit", "1");
            writer.WriteAttributeString("topLeftCell", "A2");
            writer.WriteAttributeString("activePane", "bottomLeft");
            writer.WriteAttributeString("state", "frozen");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Column widths
            if (sheet.ColumnWidths.Count > 0)
            {
                writer.WriteStartElement("cols");
                foreach (var (index, width) in sheet.ColumnWidths)
                {
                    writer.WriteStartElement("col");
                    writer.WriteAttributeString("min", (index + 1).ToString());
                    writer.WriteAttributeString("max", (index + 1).ToString());
                    writer.WriteAttributeString("width", width.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("customWidth", "1");
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            // Sheet data
            writer.WriteStartElement("sheetData");
            for (var rowIdx = 0; rowIdx < sheet.Rows.Count; rowIdx++)
            {
                var row = sheet.Rows[rowIdx];
                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", (rowIdx + 1).ToString());
                for (var colIdx = 0; colIdx < row.Count; colIdx++)
                {
                    var cell = row[colIdx];
                    if (cell.Type == CellType.Empty)
                    {
                        continue;
                    }

                    var cellRef = GetColumnLetter(colIdx) + (rowIdx + 1);
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);

                    if (cell.StyleIndex > 0)
                    {
                        writer.WriteAttributeString("s", cell.StyleIndex.ToString());
                    }

                    switch (cell.Type)
                    {
                        case CellType.InlineString:
                            writer.WriteAttributeString("t", "inlineStr");
                            writer.WriteStartElement("is");
                            writer.WriteRaw(cell.Value!);
                            writer.WriteEndElement();
                            break;
                        case CellType.Number:
                            writer.WriteStartElement("v");
                            writer.WriteString(cell.Value!);
                            writer.WriteEndElement();
                            break;
                        case CellType.Boolean:
                            writer.WriteAttributeString("t", "b");
                            writer.WriteStartElement("v");
                            writer.WriteString(cell.Value!);
                            writer.WriteEndElement();
                            break;
                    }

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            // AutoFilter
            if (sheet.FilterFirstColumn >= 0)
            {
                var firstCol = GetColumnLetter(sheet.FilterFirstColumn);
                var lastCol = GetColumnLetter(sheet.FilterLastColumn);
                writer.WriteStartElement("autoFilter");
                writer.WriteAttributeString("ref", $"{firstCol}1:{lastCol}{sheet.Rows.Count}");
                writer.WriteEndElement();
            }

            // Hyperlinks
            if (sheet.Hyperlinks.Count > 0)
            {
                writer.WriteStartElement("hyperlinks");
                for (var i = 0; i < sheet.Hyperlinks.Count; i++)
                {
                    var link = sheet.Hyperlinks[i];
                    writer.WriteStartElement("hyperlink");
                    writer.WriteAttributeString("ref", link.CellReference);
                    writer.WriteAttributeString("r", "id", NsRelationships, $"rId{i + 1}");
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        });

    static void WriteSheetRels(ZipArchive archive, SheetData sheet, int sheetIndex) =>
        WriteEntry(archive, $"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", NsPackageRelationships);
            for (var i = 0; i < sheet.Hyperlinks.Count; i++)
            {
                var link = sheet.Hyperlinks[i];
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", $"rId{i + 1}");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                writer.WriteAttributeString("Target", link.Url);
                writer.WriteAttributeString("TargetMode", "External");
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        });

    static void WriteDefault(XmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default");
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    static void WriteOverride(XmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override");
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    static readonly DateTimeOffset epoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    static void WriteEntry(ZipArchive archive, string entryPath, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(entryPath, CompressionLevel.NoCompression);
        entry.LastWriteTime = epoch;
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, xmlSettings);
        write(writer);
    }

    internal static string GetColumnLetter(int columnIndex)
    {
        var result = "";
        var index = columnIndex;
        while (index >= 0)
        {
            result = (char)('A' + index % 26) + result;
            index = index / 26 - 1;
        }

        return result;
    }
}
