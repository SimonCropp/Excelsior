namespace Excelsior;

public static class GeneratedActivators
{
    static class Holder<T>
    {
        public static Func<IReadOnlyDictionary<string, object?>, T>? Factory;
    }

    public static void Register<T>(Func<IReadOnlyDictionary<string, object?>, T> factory) =>
        Holder<T>.Factory = factory;

    internal static Func<IReadOnlyDictionary<string, object?>, T>? TryGet<T>() =>
        Holder<T>.Factory;
}
