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
    static Func<Enum, string> enumRender = Extensions.DisplayName;
    static Dictionary<Type, string> nullDisplay = [];

    public static void ForEnums(Func<Enum, string> func)
    {
        ThrowIfBookBuilderUsed();

        enumRender = func;
    }

    internal static string RenderEnum(Enum value) =>
        enumRender(value);

    public static void For<T>(Func<T, string> func)
        where T : notnull
    {
        ThrowIfBookBuilderUsed();

        renders[typeof(T)] = _ => func((T)_);
    }

    public static void NullDisplayFor<T>(string value)
        where T : notnull
    {
        ThrowIfBookBuilderUsed();

        nullDisplay[typeof(T)] = value;
    }

    static void ThrowIfBookBuilderUsed()
    {
        if (bookBuilderUsed)
        {
            throw new("Any calls to ValueRenderer must be done before any calls to BookBuilder. The recommended approach is to place calls to ValueRenderer in a ModuleInitializer.");
        }
    }

    internal static Func<object, string>? GetRender(Type type)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsAssignableTo(type))
            {
                return value;
            }
        }

        if (type.IsEnum)
        {
            //TODO: should cache this
            return _ => enumRender((Enum)_);
        }

        return null;
    }

    internal static string? GetNullDisplay(Type type)
    {
        foreach (var (key, value) in nullDisplay)
        {
            if (key.IsAssignableTo(type))
            {
                return value;
            }
        }

        return null;
    }

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}