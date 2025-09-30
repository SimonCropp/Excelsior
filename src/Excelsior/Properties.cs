static class Properties<T>
{
    static Properties() =>
        Items = typeof(T)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(_ => _.CanRead)
            .Select(_ => new Property<T>(_))
            .ToList();

    public static IReadOnlyList<Property<T>> Items { get; }
}