static class Properties<T>
{
    static Properties() =>
        Items = GetPropertiesRecursive(typeof(T), [], false).ToList();

    static IEnumerable<Property<T>> GetPropertiesRecursive(Type type, Stack<(PropertyInfo property, ParameterInfo? parameter)> stack, bool useHierachyForName)
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

            stack.Push((property, parameter));

            var infos = stack.Reverse().ToList();
            var func = CreateGet(infos.Select(_ => _.property));
            if (ShouldSplit(property, parameter, out var nestedUseHierachyForName))
            {
                foreach (var nested in GetPropertiesRecursive(property.PropertyType, stack, nestedUseHierachyForName))
                {
                    yield return nested;
                }
            }
            else
            {
                yield return new(property, parameter, func, infos, useHierachyForName);
            }

            stack.Pop();
        }
    }

    static bool ShouldSplit(PropertyInfo property, ParameterInfo? parameter, out bool useHierachyForName)
    {
        var split = property.Attribute<SplitAttribute>() ??
                    parameter?.Attribute<SplitAttribute>() ??
                    property.PropertyType.Attribute<SplitAttribute>();
        if (split == null)
        {
            useHierachyForName = false;
            return false;
        }

        useHierachyForName = split.UseHierachyForName;
        return true;
    }

    static ParameterExpression targetParam = Expression.Parameter(typeof(T));

    static Func<T, object?> CreateGet(IEnumerable<PropertyInfo> path)
    {
        Expression current = targetParam;

        foreach (var property in path)
        {
            var propertyAccess = Expression.Property(current, property);

            if (current.Type.IsValueType &&
                Nullable.GetUnderlyingType(current.Type) == null)
            {
                current = propertyAccess;
                continue;
            }

            // Add null check if the current type is nullable (reference type or Nullable<T>)
            // current != null ? current.Property : default(PropertyType)
            current = Expression.Condition(
                Expression.NotEqual(current, Expression.Constant(null, current.Type)),
                propertyAccess,
                Expression.Default(property.PropertyType),
                property.PropertyType
            );
        }

        var box = Expression.Convert(current, typeof(object));
        return Expression.Lambda<Func<T, object?>>(box, targetParam).Compile();
    }

    public static IReadOnlyList<Property<T>> Items { get; }
}