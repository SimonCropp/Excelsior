namespace Excelsior;

[AttributeUsage(AttributeTargets.Property)]
public sealed class ColumnAttribute : Attribute
{
    public string? HeaderText { get; set; }
    public int Order { get; set; } = -1;
    public double Width { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplayText { get; set; }
    public bool TreatAsHtml { get; set; }
}