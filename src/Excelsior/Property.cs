class Property<T>(PropertyInfo info)
{
    static int? GetOrder(PropertyInfo info)
    {
        var attribute = info.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Order;
    }

    public Func<T, object?> Get { get; } = CreateGet(info);
    public string DisplayName { get; } = GetDisplayName(info);
    public string Name { get; } = info.Name;
    public int? Order { get; } = GetOrder(info);
    public Type Type { get; } = info.PropertyType;

    static string GetDisplayName(PropertyInfo info)
    {
        var display = info.GetCustomAttribute<DisplayAttribute>();
        if (display?.Name != null)
        {
            return display.Name;
        }

        var displayName = info.GetCustomAttribute<DisplayNameAttribute>();
        if (displayName != null)
        {
            return displayName.DisplayName;
        }

        return CamelCase.Split(info.Name);
    }

    static ParameterExpression targetParam = Expression.Parameter(typeof(T));

    static Func<T, object?> CreateGet(PropertyInfo info)
    {
        var property = Expression.Property(targetParam, info);
        var box = Expression.Convert(property, typeof(object));
        return Expression.Lambda<Func<T, object?>>(box, targetParam).Compile();
    }
}