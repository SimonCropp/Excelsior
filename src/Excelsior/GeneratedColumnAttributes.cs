namespace Excelsior;

public static class GeneratedColumnAttributes
{
    static ConcurrentDictionary<(Type type, string property), Values> store = new();

    public static void Register(Type type, string propertyName, Values values) =>
        store[(type, propertyName)] = values;

    internal static Values? TryGet(Type type, string propertyName) =>
        store.GetValueOrDefault((type, propertyName));

    public record Values(
        string? Heading = null,
        int? Order = null,
        int? Width = null,
        string? Format = null,
        string? NullDisplay = null,
        bool IsHtml = false,
        bool? Filter = null,
        bool? Include = null);
}
