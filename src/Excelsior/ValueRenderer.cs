namespace Excelsior;

public static partial class ValueRenderer
{
    internal static bool TrimWhitespace { get; private set; } = true;
    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders =[];
    static Dictionary<Type, Func<object, string>> itemRenders = [];
    static Func<Enum, string> enumRender = EnumExtensions.Humanize;
    static ConcurrentDictionary<Type, (bool isEnumerable, Func<object, string>? render)> renderCache = [];

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

        renders[typeof(T)] = _ => func((T) _);
        itemRenders[typeof(T)] = _ => func((T) _);
    }

    static void ThrowIfBookBuilderUsed()
    {
        if (bookBuilderUsed)
        {
            throw new("Any calls to ValueRenderer must be done before any calls to BookBuilder. The recommended approach is to place calls to ValueRenderer in a ModuleInitializer.");
        }
    }

    static bool IsTypeCompatible(Type type, Type key) =>
        type.IsAssignableTo(key) ||
        Nullable.GetUnderlyingType(type)?.IsAssignableTo(key) == true;

    internal static (bool isEnumerable, Func<object, string>? render) GetRender(Type type) =>
        renderCache.GetOrAdd(type, ResolveRender);

    static TValue? FindBestMatch<TValue>(Type type, Dictionary<Type, TValue> dictionary)
        where TValue : class
    {
        TValue? result = null;
        var underlyingType = Nullable.GetUnderlyingType(type) ?? type;

        foreach (var (key, value) in dictionary)
        {
            if (!IsTypeCompatible(type, key))
            {
                continue;
            }

            if (key == underlyingType)
            {
                return value;
            }

            result ??= value;
        }

        return result;
    }

    static (bool isEnumerable, Func<object, string>? render) ResolveRender(Type type)
    {
        var render = FindBestMatch(type, renders);
        if (render != null)
        {
            return (false, render);
        }

        if (type.IsAssignableTo<IEnumerable<string>>())
        {
            return (true, null);
        }

        if (type.IsAssignableTo<IEnumerable<Link>>())
        {
            return (true, null);
        }

        foreach (var enumerableType in GetEnumerableTypes(type))
        {
            var itemRender = FindBestMatch(enumerableType, itemRenders);
            if (itemRender != null)
            {
                return (true, itemRender);
            }
        }

        if (type.IsEnum ||
            Nullable.GetUnderlyingType(type)?.IsEnum == true)
        {
            return (false, _ => enumRender((Enum)_));
        }

        return (false, null);
    }

    static IEnumerable<Type> GetEnumerableTypes(Type type)
    {
        if (type == typeof(string))
        {
            yield break;
        }

        if (TryGetEnumerableType(type, out var enumerableType))
        {
            yield return enumerableType;
        }

        foreach (var @interface in type.GetInterfaces())
        {
            if (TryGetEnumerableType(@interface, out enumerableType))
            {
                yield return enumerableType;
            }
        }
    }

    static bool TryGetEnumerableType(Type interfacteType, [NotNullWhen(true)] out Type? enumerableType)
    {
        if (interfacteType.IsGenericType &&
            interfacteType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
        {
            enumerableType = interfacteType.GetGenericArguments()[0];
            return true;
        }

        enumerableType = null;
        return false;
    }

    internal static void Reset()
    {
        bookBuilderUsed = false;
        TrimWhitespace = true;
        renders = [];
        itemRenders = [];
        enumRender = EnumExtensions.Humanize;
        renderCache = [];
        nullDisplay = [];
        DefaultDateFormat = "yyyy-MM-dd";
        DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";
        DefaultDateTimeOffsetFormat = "yyyy-MM-dd HH:mm:ss z";
    }

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}