namespace Excelsior;

public static partial class ValueRenderer
{
    [StringSyntax(StringSyntaxAttribute.DateOnlyFormat)]
    public static string DefaultDateFormat
    {
        internal get;
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
    } = "yyyy-MM-dd";

    [StringSyntax(StringSyntaxAttribute.DateTimeFormat)]
    public static string DefaultDateTimeFormat
    {
        internal get;
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
    } = "yyyy-MM-dd HH:mm:ss";

    [StringSyntax(StringSyntaxAttribute.DateTimeFormat)]
    public static string DefaultDateTimeOffsetFormat
    {
        internal get;
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
    } = "yyyy-MM-dd HH:mm:ss z";
}
