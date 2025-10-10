static class Properties<T>
{
    static Properties() =>
        Items = GetPropertiesRecursive(typeof(T), []).ToList();

    static IEnumerable<Property<T>> GetPropertiesRecursive(Type type, Stack<PropertyInfo> stack)
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

            stack.Push(property);

            var infos = stack.Reverse().ToList();
            var func = CreateGet(infos);
            if (ShouldSplit(property, parameter))
            {
                yield return new(property, parameter, func, infos);
            }
            else
            {
                foreach (var nested in GetPropertiesRecursive(property.PropertyType, stack))
                {
                    yield return nested;
                }
            }

            stack.Pop();
        }
    }

    static bool ShouldSplit(PropertyInfo property, ParameterInfo? parameter)
    {
        var split = property.Attribute<SplitAttribute>() ??
                    parameter?.Attribute<SplitAttribute>() ??
                    property.PropertyType.Attribute<SplitAttribute>();
        return split == null;
    }

    static ParameterExpression targetParam = Expression.Parameter(typeof(T));

    static Func<T, object?> CreateGet(IReadOnlyList<PropertyInfo> path)
    {
        Expression current = targetParam;
        foreach (var property in path)
        {
            current = Expression.Property(current, property);
        }

        var box = Expression.Convert(current, typeof(object));
        return Expression.Lambda<Func<T, object?>>(box, targetParam).Compile();
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}