namespace Excelsior;

public static class ValueRenderer
{
    public static string DefaultDateFormat
    {
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
        internal get;
    } = "yyyy-MM-dd";

    public static string DefaultDateTimeFormat
    {
        set
        {
            ThrowIfBookBuilderUsed();
            field = value;
        }
        internal get;
    } = "yyyy-MM-dd HH:mm:ss";

    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders = [];

    public static void For<T>(Func<T, string> func)
    {
        ThrowIfBookBuilderUsed();

        renders[typeof(T)] = _ => func((T) _);
    }

    static void ThrowIfBookBuilderUsed()
    {
        if (bookBuilderUsed)
        {
            throw new("Any calls to ValueRenderer must be done before any calls to BookBuilder. The recommended approach is to place calls to ValueRenderer in a ModuleInitializer.");
        }
    }

    internal static Func<object, string>? GetRender(Type memberType)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsAssignableTo(memberType))
            {
                return value;
            }
        }

        return null;
    }

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}