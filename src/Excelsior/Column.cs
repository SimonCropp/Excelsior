namespace Excelsior;

class Column<TStyle, TModel>
{
    public required string Heading { get; set; }
    public required int? Order { get; set; }
    public required int? Width { get; set; }
    public required Action<TStyle>? HeadingStyle { get; set; }
    public required Action<TStyle, TModel, object?>? CellStyle { get; set; }
    public required string? Format { get; set; }
    public required string? NullDisplay { get; set; }
    public required Func<TModel, object, string?>? Render { get; set; }
    public required bool IsHtml { get; set; }
    public required bool IsNumber { get; init; }
    public required string Name { get; set; }
    public required Func<TModel, object?> GetValue { get; init; }
}

public class Column<TStyle, TModel, TProperty>
{
    public string? Heading { get; set; }
    public int? Order { get; set; }
    public int? Width { get; set; }
    public Action<TStyle>? HeadingStyle { get; set; }
    public Action<TStyle, TModel, TProperty>? CellStyle { get; set; }
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public Func<TModel, TProperty, string?>? Render { get; set; }
    public bool? IsHtml { get; set; }
}