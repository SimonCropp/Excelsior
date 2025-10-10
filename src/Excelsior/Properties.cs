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

            var func = CreateGet(property);
            var split = property.Attribute<SplitAttribute>() ?? parameter?.Attribute<SplitAttribute>();
            if (split == null)
            {
                yield return new(property, parameter, null);
            }

        }
    }
    static ParameterExpression targetParam = Expression.Parameter(typeof(T));

    static Func<T, object?> CreateGet(PropertyInfo info)
    {
        var property = Expression.Property(targetParam, info);
        var box = Expression.Convert(property, typeof(object));
        return Expression.Lambda<Func<T, object?>>(box, targetParam).Compile();
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}