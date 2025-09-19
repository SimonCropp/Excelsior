namespace Excelsior;

class ColumnSettings
{
    public required string? HeaderText { get; init; }
    public required int? Order { get; init; }
    public required double? ColumnWidth { get; init; }
    public required Action<IXLStyle>? HeaderStyle { get; init; }
    public required Action<IXLStyle>? DataCellStyle { get; init; }
    public required Action<IXLStyle, object?>? ConditionalStyling { get; init; }
    public required string? DateTimeFormat { get; init; }
    public required string? NumberFormat { get; init; }
    public required string? NullDisplayText { get; init; }
    public required Func<object, string>? CustomFormatter { get; init; }
    public required Func<bool, string>? BooleanDisplayFormat { get; init; }
    public required Func<Enum, string>? EnumDisplayFormat { get; init; }
}

public class ColumnSettings<T>
{
    public string? HeaderText { get; set; }
    public int? Order { get; set; }
    public double? ColumnWidth { get; set; }
    public Action<IXLStyle>? HeaderStyle { get; set; }
    public Action<IXLStyle>? DataCellStyle { get; set; }
    public Action<IXLStyle, T>? ConditionalStyling { get; set; }
    public string? DateTimeFormat { get; set; }
    public string? NumberFormat { get; set; }
    public string? NullDisplayText { get; set; }
    public Func<object, string>? CustomFormatter { get; set; }
    public Func<bool, string>? BooleanDisplayFormat { get; set; }
    public Func<Enum, string>? EnumDisplayFormat { get; set; }
}