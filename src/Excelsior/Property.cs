class Property<T>
{
    public Property(
        PropertyInfo info,
        ParameterInfo? constructorParameter,
        Func<T, object?> get,
        IReadOnlyList<PropertyInfo> infos,
        bool useHierachyForName)
    {
        Get = get;
        var column = info.Attribute<ColumnAttribute>() ?? constructorParameter?.Attribute<ColumnAttribute>();
        var display = info.Attribute<DisplayAttribute>() ?? constructorParameter?.Attribute<DisplayAttribute>();
        var displayName = info.Attribute<DisplayNameAttribute>() ?? constructorParameter?.Attribute<DisplayNameAttribute>();
        DisplayName = GetHeading(info, display, column, displayName);
        Name = string.Join('.', infos.Select(_ => _.Name));
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

    static int? GetWidth(ColumnAttribute? column)
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
    public int? Width { get; }
    public string? Format { get; }
    public string? NullDisplay { get; }
    public bool IsHtml { get; }

    static string GetHeading(PropertyInfo info, DisplayAttribute? display, ColumnAttribute? column, DisplayNameAttribute? displayName)
    {
        if (column?.Heading != null)
        {
            return column.Heading;
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
}