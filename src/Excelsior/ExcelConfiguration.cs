/// <summary>
/// Global Excel configuration
/// </summary>
public class ExcelConfiguration
{
    public string WorksheetName { get; set; } = "Sheet1";
    public bool AutoSizeColumns { get; set; } = true;
    public bool UseAlternatingRowColors { get; set; }
    public XLColor AlternateRowColor { get; set; } = XLColor.LightGray;
    public Action<IXLStyle>? HeaderStyle { get; set; }
    public Action<IXLStyle>? GlobalStyle { get; set; }
}