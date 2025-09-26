namespace Excelsior;

public static class ValueRenderer
{
    static Dictionary<Type, Func<object, string>> renders = [];

    public static void For<T>(Func<T, string> func) =>
        renders[typeof(T)] = _ => func((T) _);

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
}