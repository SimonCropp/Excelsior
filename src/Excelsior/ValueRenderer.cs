namespace Excelsior;

public static class ValueRenderer
{
    private static bool bookBuilderUsed;
    static Dictionary<Type, Func<object, string>> renders = [];

    public static void For<T>(Func<T, string> func)
    {
        if (bookBuilderUsed)
        {
            throw new("Any calls to ValueRenderer.For must be done before any calls to BookBuilder. The recommended approach is to place calls to ValueRenderer.For in a ModuleInitializer.");
        }

        renders[typeof(T)] = _ => func((T) _);
    }

    internal static bool TryRender(Type memberType, object instance, [NotNullWhen(true)] out string? result)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsAssignableTo(memberType))
            {
                result = value(instance);
                return true;
            }
        }

        result = null;
        return false;
    }

    internal static void SetBookBuilderUsed() => bookBuilderUsed = true;
}