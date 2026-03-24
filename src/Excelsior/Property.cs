class Property<T>
{
    public Property(
        PropertyInfo info,
        ParameterInfo? constructorParameter,
        Func<T, object?> get,
        IReadOnlyList<(PropertyInfo property, ParameterInfo? parameter)> infos,
        bool useHierachyForName)
    {
        Get = get;
        var generated = GeneratedColumnAttributes.TryGet(info.DeclaringType!, info.Name);
        var display = info.Attribute<DisplayAttribute>() ?? constructorParameter?.Attribute<DisplayAttribute>();
        DisplayName = GetHeading(infos, useHierachyForName);
        Name = string.Join('.', infos.Select(_ => _.property.Name));

        if (generated is null)
        {
            var column = info.Attribute<ColumnAttribute>() ?? constructorParameter?.Attribute<ColumnAttribute>();
            Order = GetOrder(column, display);
            Width = GetWidth(column);
            Format = column?.Format;
            NullDisplay = column?.NullDisplay;
            IsHtml = column?.IsHtml ?? false;
            Filter = column is {FilterHasValue: true} ? column.Filter : null;
            Include = column is {IncludeHasValue: true} ? column.Include : null;
        }
        else
        {
            Order = generated.Order ?? display?.Order;
            Width = generated.Width;
            Format = generated.Format;
            NullDisplay = generated.NullDisplay;
            IsHtml = generated.IsHtml;
            Filter = generated.Filter;
            Include = generated.Include;
        }

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
    public bool? Filter { get; }
    public bool? Include { get; }

    static string GetHeading(IReadOnlyList<(PropertyInfo property, ParameterInfo? parameter)> infos, bool useHierachyForName)
    {
        var names = new List<string>();
        foreach (var (property, parameter) in infos)
        {
            Add(property, parameter);
        }

        if (!useHierachyForName)
        {
            return names.Last();
        }

        return string.Join(' ', names);

        void Add(PropertyInfo property, ParameterInfo? parameter)
        {
            var generated = GeneratedColumnAttributes.TryGet(property.DeclaringType!, property.Name);
            if (generated?.Heading is not null)
            {
                names.Add(generated.Heading);
                return;
            }

            if (generated is null)
            {
                var column = property.Attribute<ColumnAttribute>() ?? parameter?.Attribute<ColumnAttribute>();
                if (column?.Heading != null)
                {
                    names.Add(column.Heading);
                    return;
                }
            }

            var display = property.Attribute<DisplayAttribute>() ?? parameter?.Attribute<DisplayAttribute>();
            if (display?.Name != null)
            {
                names.Add(display.Name);
                return;
            }

            var displayName = property.Attribute<DisplayNameAttribute>() ?? parameter?.Attribute<DisplayNameAttribute>();
            if (displayName != null)
            {
                names.Add(displayName.DisplayName);
                return;
            }

            names.Add(CamelCase.Split(property.Name));
        }
    }
}