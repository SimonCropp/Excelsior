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

            if (property.Attribute<IgnoreAttribute>() != null)
            {
                continue;
            }

            var parameter = defaultConstructor?.SingleOrDefault(_ => _.Name == property.Name);
            if (parameter.Attribute<IgnoreAttribute>() != null)
            {
                continue;
            }

            var split = property.Attribute<SplitAttribute>() ?? parameter?.Attribute<SplitAttribute>();
            if (split == null)
            {
                yield return new(property, parameter, null);
            }

        }
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}