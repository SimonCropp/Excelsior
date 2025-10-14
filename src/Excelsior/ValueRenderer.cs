namespace Excelsior;

public static partial class ValueRenderer
{
    internal static bool TrimWhitespace { get; private set; } = true;
    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders =[];
    static Dictionary<Type, Func<object, string>> enumerableRenders = [];
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

        renders[typeof(T)] = _ => func((T) _);
        enumerableRenders[typeof(T)] = _ =>
        {
            var enumerable = (IEnumerable<T?>) _;
            return ListBuilder.Build(enumerable.Select(_ => _ == null ? null : func(_)));
        };
    }

    static void ThrowIfBookBuilderUsed()
    {
        if (bookBuilderUsed)
        {
            throw new("Any calls to ValueRenderer must be done before any calls to BookBuilder. The recommended approach is to place calls to ValueRenderer in a ModuleInitializer.");
        }
    }

    internal static (bool isEnumerable, Func<object, string>? render) GetRender(Type type)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsAssignableTo(type))
            {
                return (false, value);
            }
        }

        if (type.IsAssignableTo<IEnumerable<string>>())
        {
            return (true, _ => ListBuilder.Build((IEnumerable<string>)_));
        }

        var enumerableTypes = GetEnumerableTypes(type).ToList();
        foreach (var enumerableType in enumerableTypes)
        {
            foreach (var (key, value) in enumerableRenders)
            {
                if (key.IsAssignableTo(enumerableType))
                {
                    return (true, value);
                }
            }
        }

        if (type.IsEnum)
        {
            //TODO: should cache this
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