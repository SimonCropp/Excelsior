namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ColumnAttribute :
    Attribute
{
    public string? Heading { get; set; }
    public int Order { get; set; } = -1;
    public int Width { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public bool IsHtml { get; set; }

    bool filter;

    public bool Filter
    {
        get => filter;
        set
        {
            filter = value;
            FilterHasValue = true;
        }
    }

    internal bool FilterHasValue { get; private set; }
}
