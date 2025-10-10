static class Properties<T>
{
    static Properties() =>
        Items = GetPropertiesRecursive(typeof(T), "").ToList();

    static IEnumerable<Property<T>> GetPropertiesRecursive(Type type, string prefix)
    {
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
            var path = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";
            var func = CreateGet(path);
            var split = property.Attribute<SplitAttribute>() ?? parameter?.Attribute<SplitAttribute>();
            if (split == null)
            {
                yield return new(property, parameter, func);
            }
            else
            {
                foreach (var nested in GetPropertiesRecursive(property.PropertyType, path))
                {
                    yield return nested;
                }
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

    static Func<T, object?> CreateGet(string propertyPath)
    {
        Expression current = targetParam;

        foreach (var propName in propertyPath.Split('.'))
        {
            var property = current.Type.GetProperty(propName);
            if (property == null)
                throw new ArgumentException($"Property '{propName}' not found");

            current = Expression.Property(current, property);
        }

        var box = Expression.Convert(current, typeof(object));
        return Expression.Lambda<Func<T, object?>>(box, targetParam).Compile();
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}