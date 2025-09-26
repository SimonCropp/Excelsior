namespace ExcelsiorAspose;

class ColumnSettings<TStyle>
{
    public required string? HeaderText { get; init; }
    public required int? Order { get; init; }
    public required double? ColumnWidth { get; init; }
    public required Action<TStyle>? HeaderStyle { get; init; }
    public required Action<TStyle>? DataCellStyle { get; init; }
    public required Action<TStyle, object?>? ConditionalStyling { get; init; }
    public required string? Format { get; init; }
    public required string? NullDisplayText { get; init; }
    public required Func<object, string?>? Render { get; init; }
}

public class ColumnSettings<TStyle, TProperty>
{
    public string? HeaderText { get; set; }
    public int? Order { get; set; }
    public double? ColumnWidth { get; set; }
    public Action<TStyle>? HeaderStyle { get; set; }
    public Action<TStyle>? DataCellStyle { get; set; }
    public Action<TStyle, TProperty>? ConditionalStyling { get; set; }
    public string? Format { get; set; }
    public string? NullDisplayText { get; set; }
    public Func<TProperty, string?>? Render { get; set; }
}