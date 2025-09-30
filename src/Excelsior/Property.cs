class Property<T>
{
    public Property(PropertyInfo info, ParameterInfo? constructorParameter)
    {
        Get = CreateGet(info);
        var column = info.GetCustomAttribute<ColumnAttribute>() ?? constructorParameter?.GetCustomAttribute<ColumnAttribute>();
        var display = info.GetCustomAttribute<DisplayAttribute>() ?? constructorParameter?.GetCustomAttribute<DisplayAttribute>();
        var displayName = info.GetCustomAttribute<DisplayNameAttribute>() ?? constructorParameter?.GetCustomAttribute<DisplayNameAttribute>();
        DisplayName = GetHeader(info, display, column,displayName);
        Name = info.Name;
        Order = GetOrder(column, display);
        Width = GetWidth(column);
        Format = column?.Format;
        NullDisplay = column?.NullDisplay;
        IsHtml = column?.IsHtml ?? false;
        Type = info.PropertyType;
        IsNumber = info.PropertyType.IsNumericType();
    }

    static int? GetOrder(ColumnAttribute? column, DisplayAttribute? display)
    {
        if (column is { Order: > -1 })
        {
            return column.Order;
        }

        return display?.Order;
    }

    static double? GetWidth(ColumnAttribute? column)
    {
        if (column is { Width: > -1 })
        {
            return column.Width;
        }

        return null;
    }

    public Func<T, object?> Get { get; }
    public string DisplayName { get; }
    public string Name { get; }
    public int? Order { get; }
    public Type Type { get; }
    public bool IsNumber { get; }
    public double? Width { get; }
    public string? Format { get; }
    public string? NullDisplay { get; }
    public bool IsHtml { get; }

    static string GetHeader(PropertyInfo info, DisplayAttribute? display, ColumnAttribute? column, DisplayNameAttribute? displayName)
    {
        if (column?.Header != null)
        {
            return column.Header;
        }

        if (display?.Name != null)
        {
            return display.Name;
        }

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