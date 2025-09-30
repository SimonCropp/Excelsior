namespace Excelsior;

class Column<TStyle, TModel>
{
    public required string Header { get; set; }
    public required int? Order { get; set; }
    public required double? Width { get; set; }
    public required Action<TStyle>? HeaderStyle { get; set; }
    public required Action<TStyle, object?>? CellStyle { get; set; }
    public required string? Format { get; set; }
    public required string? NullDisplay { get; set; }
    public required Func<object, string?>? Render { get; set; }
    public required bool IsHtml { get; set; }
    public required bool IsNumber { get; set; }
    public required string Name { get; set; }
    public required Func<object, object?> GetValue { get; set; }
}

public class Column<TStyle, TModel, TProperty>
{
    public string? Header { get; set; }
    public int? Order { get; set; }
    public double? Width { get; set; }
    public Action<TStyle>? HeaderStyle { get; set; }
    public Action<TStyle, TProperty>? CellStyle { get; set; }
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public Func<TProperty, string?>? Render { get; set; }
    public bool? IsHtml { get; set; }
}