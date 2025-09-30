namespace Excelsior;

class Column<TStyle>
{
    public required string HeaderText { get; set; }
    public required int? Order { get; set; }
    public required double? ColumnWidth { get; set; }
    public required Action<TStyle>? HeaderStyle { get; set; }
    public required Action<TStyle, object?>? CellStyle { get; set; }
    public required string? Format { get; set; }
    public required string? NullDisplayText { get; set; }
    public required Func<object, string?>? Render { get; set; }
    public required bool TreatAsHtml { get; set; }
    public required bool IsNumber { get; set; }
    public required string Name { get; set; }
    public required Func<object, object?> GetValue { get; set; }
}

public class Column<TStyle, TProperty>
{
    public string? HeaderText { get; set; }
    public int? Order { get; set; }
    public double? ColumnWidth { get; set; }
    public Action<TStyle>? HeaderStyle { get; set; }
    public Action<TStyle, TProperty>? CellStyle { get; set; }
    public string? Format { get; set; }
    public string? NullDisplayText { get; set; }
    public Func<TProperty, string?>? Render { get; set; }
    public bool? TreatAsHtml { get; set; }
}