/// <summary>
/// Configuration for individual columns
/// </summary>
public class ColumnConfiguration
{
    public string? HeaderText { get; set; }
    public int? Order { get; set; }
    public double? ColumnWidth { get; set; }
    public Action<IXLStyle>? HeaderStyle { get; set; }
    public Action<IXLStyle>? DataCellStyle { get; set; }
    public Action<IXLStyle, object?>? ConditionalStyling { get; set; }
    public string? DateTimeFormat { get; set; }
    public string? NumberFormat { get; set; }
    public string? NullDisplayText { get; set; }
    public Func<object, string>? CustomFormatter { get; set; }
    public Func<bool, string>? BooleanDisplayFormat { get; set; }
    public Func<Enum, string>? EnumDisplayFormat { get; set; }
}