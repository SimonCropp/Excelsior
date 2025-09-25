class ColumnSettings
{
    public required string? HeaderText { get; init; }
    public required int? Order { get; init; }
    public required double? ColumnWidth { get; init; }
    public required Action<Style>? HeaderStyle { get; init; }
    public required Action<Style>? DataCellStyle { get; init; }
    public required Action<Style, object?>? ConditionalStyling { get; init; }
    public required string? Format { get; init; }
    public required string? NullDisplayText { get; init; }
    public required Func<object, string?>? Render { get; init; }
}

public class ColumnSettings<T>
{
    public string? HeaderText { get; set; }
    public int? Order { get; set; }
    public double? ColumnWidth { get; set; }
    public Action<Style>? HeaderStyle { get; set; }
    public Action<Style>? DataCellStyle { get; set; }
    public Action<Style, T>? ConditionalStyling { get; set; }
    public string? Format { get; set; }
    public string? NullDisplayText { get; set; }
    public Func<T, string?>? Render { get; set; }
}