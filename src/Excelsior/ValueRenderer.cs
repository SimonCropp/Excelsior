namespace Excelsior;

public static partial class ValueRenderer
{
    internal static bool TrimWhitespace { get; private set; } = true;
    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders = [];
    static Func<Enum, string> enumRender = Extensions.DisplayName;

    public static void DisableWhitespaceTrimming()
    {
        ThrowIfBookBuilderUsed();

        TrimWhitespace = false;
    }

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

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}