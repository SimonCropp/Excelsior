namespace Excelsior;

public static partial class ValueRenderer
{
    internal static bool TrimWhitespace { get; private set; } = true;
    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders =[];
    static Dictionary<Type, Func<object, string>> itemRenders = [];
    static Func<Enum, string> enumRender = EnumExtensions.Humanize;
    static Dictionary<Type, (bool isEnumerable, Func<object, string>? render)> renderCache = [];

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

    internal static (bool isEnumerable, Func<object, string>? render) GetRender(Type type)
    {
        if (renderCache.TryGetValue(type, out var cached))
        {
            return cached;
        }

        var result = ResolveRender(type);
        renderCache[type] = result;
        return result;
    }

    static (bool isEnumerable, Func<object, string>? render) ResolveRender(Type type)
    {
        foreach (var (key, value) in renders)
        {
            if (IsTypeCompatible(type, key))
            {
                return (false, value);
            }
        }

        if (type.IsAssignableTo<IEnumerable<string>>())
        {
            return (true, null);
        }

        foreach (var enumerableType in GetEnumerableTypes(type))
        {
            foreach (var (key, value) in itemRenders)
            {
                if (IsTypeCompatible(enumerableType, key))
                {
                    return (true, value);
                }
            }
        }

        if (type.IsEnum)
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

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}