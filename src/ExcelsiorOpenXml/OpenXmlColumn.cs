namespace ExcelsiorOpenXml;

public class OpenXmlColumn(DocumentFormat.OpenXml.Spreadsheet.Column column)
{
    public DocumentFormat.OpenXml.Spreadsheet.Column Column { get; } = column;
    public double? Width { get; set; }
}
