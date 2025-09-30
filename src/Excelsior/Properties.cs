static class Properties<T>
{
    static Properties()
    {
        var type = typeof(T);
        var defaultConstructor = type.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Select(_ => _.GetParameters())
            .OrderBy(_ => _.Length)
            .FirstOrDefault();
        var properties = type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(_ => _.CanRead)
            .Select(property =>
            {
                var constructorParameter = defaultConstructor?.SingleOrDefault(_ => _.Name == property.Name);
                return new Property<T>(property, constructorParameter);
            })
            .ToList();

        Items = properties;
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}