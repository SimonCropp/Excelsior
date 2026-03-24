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

    internal static string? GetNullDisplay(Type type) =>
        FindBestMatch(type, nullDisplay);
}