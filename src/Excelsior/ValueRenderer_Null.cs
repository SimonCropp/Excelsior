namespace Excelsior;

public static partial class ValueRenderer
{
    static Dictionary<Type, string> nullDisplay = [];

    public static void NullDisplayFor<T>(string value)
        where T : notnull
    {
        ThrowIfBookBuilderUsed();

        nullDisplay[typeof(T)] = value;
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
}