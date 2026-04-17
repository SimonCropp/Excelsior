namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ColumnAttribute :
    Attribute
{
    public string? Heading { get; set; }
    public int Order { get; set; } = -1;
    public int Width { get; set; } = -1;
    public int MinWidth { get; set; } = -1;
    public int MaxWidth { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }

    public bool IsHtml
    {
        get;
        set
        {
            field = value;
            IsHtmlHasValue = true;
        }
    }

    internal bool IsHtmlHasValue { get; private set; }

    public bool Filter
    {
        get;
        set
        {
            field = value;
            FilterHasValue = true;
        }
    }

    internal bool FilterHasValue { get; private set; }

    public bool Include
    {
        get;
        set
        {
            field = value;
            IncludeHasValue = true;
        }
    } = true;

    internal bool IncludeHasValue { get; private set; }
}
