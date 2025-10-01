namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ColumnAttribute :
    Attribute
{
    public string? Header { get; set; }
    public int Order { get; set; } = -1;
    public double Width { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public bool IsHtml { get; set; }
}