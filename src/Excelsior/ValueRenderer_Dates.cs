namespace Excelsior;

public static partial class ValueRenderer
{
    [StringSyntax(StringSyntaxAttribute.DateOnlyFormat)]
    public static string DefaultDateFormat
    {
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
        internal get;
    } = "yyyy-MM-dd";

    [StringSyntax(StringSyntaxAttribute.DateTimeFormat)]
    public static string DefaultDateTimeFormat
    {
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
        internal get;
    } = "yyyy-MM-dd HH:mm:ss";

    [StringSyntax(StringSyntaxAttribute.DateTimeFormat)]
    public static string DefaultDateTimeOffsetFormat
    {
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
        internal get;
    } = "yyyy-MM-dd HH:mm:ss z";
}