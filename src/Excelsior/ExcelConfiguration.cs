namespace Excelsior;

/// <summary>
/// Global Excel configuration
/// </summary>
public class ExcelConfiguration
{
    public string WorksheetName { get; set; } = "Sheet1";
    public bool UseAlternatingRowColors { get; set; }
    public XLColor AlternateRowColor { get; set; } = XLColor.LightGray;
    public Action<IXLStyle>? GlobalStyle { get; set; }
}