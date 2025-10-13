namespace Excelsior;

public static partial class ValueRenderer
{
    static ValueRenderer()
    {
        enumerableRenders = [];
        enumerableRenders[typeof(string)] = _ => ListBuilder.Build((IEnumerable<string>)_);
    }

    internal static bool TrimWhitespace { get; private set; } = true;
    static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders =[];
    static Dictionary<Type, Func<object, string>> enumerableRenders;
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

        foreach (var enumerableType in GetEnumerableTypes(type))
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
            return [];
        }

        return type
            .GetInterfaces()
            .Where(i => i.IsGenericType &&
                        i.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            .Select(i => i.GetGenericArguments()[0]);
    }

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}