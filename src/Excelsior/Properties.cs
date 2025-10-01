static class Properties<T>
{
    static Properties() =>
        Items = GetProperties().ToList();

    static IEnumerable<Property<T>> GetProperties()
    {
        var type = typeof(T);
        var defaultConstructor = type.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Select(_ => _.GetParameters())
            .OrderByDescending(_ => _.Length)
            .FirstOrDefault();
        foreach (var property in type
                     .GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanRead)
            {
                continue;
            }

            if (property.GetCustomAttribute<IgnoreAttribute>() != null)
            {
                continue;
            }

            var constructorParameter = defaultConstructor?.SingleOrDefault(_ => _.Name == property.Name);
            if (constructorParameter?.GetCustomAttribute<IgnoreAttribute>() != null)
            {
                continue;
            }

            yield return new(property, constructorParameter);
        }
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}