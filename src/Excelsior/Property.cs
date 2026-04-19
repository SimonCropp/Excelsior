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
            MinWidth = GetMinWidth(column);
            MaxWidth = GetMaxWidth(column);
            Format = column?.Format;
            NullDisplay = column?.NullDisplay;
            (IsHtml, IsHtmlExplicit) = ResolveIsHtml(info, constructorParameter, column);
            Filter = column is {FilterHasValue: true} ? column.Filter : null;
            Include = column is {IncludeHasValue: true} ? column.Include : null;
        }
        else
        {
            Order = generated.Order ?? display?.Order;
            Width = generated.Width;
            MinWidth = generated.MinWidth;
            MaxWidth = generated.MaxWidth;
            Format = generated.Format;
            NullDisplay = generated.NullDisplay;
            IsHtml = generated.IsHtml;
            IsHtmlExplicit = generated.IsHtml;
            Filter = generated.Filter;
            Include = generated.Include;
        }

        Type = info.PropertyType;
        IsNumber = info.PropertyType.IsNumericType();
    }

    static (bool isHtml, bool isExplicit) ResolveIsHtml(PropertyInfo info, ParameterInfo? constructorParameter, ColumnAttribute? column)
    {
        var columnExplicit = column is {IsHtmlHasValue: true} ? column.IsHtml : (bool?)null;
        var hasHtmlSyntax = HasHtmlStringSyntax(info, constructorParameter);

        if (columnExplicit == false && hasHtmlSyntax)
        {
            throw new($"Property '{info.DeclaringType!.Name}.{info.Name}': mismatched IsHtml — [Column(IsHtml = false)] contradicts [StringSyntax(\"html\")].");
        }

        if (columnExplicit is { } explicitValue)
        {
            return (explicitValue, true);
        }

        if (hasHtmlSyntax)
        {
            return (true, true);
        }

        if (HasHtmlAttribute(info, constructorParameter))
        {
            return (true, true);
        }

        return (false, false);
    }

    static bool HasHtmlStringSyntax(PropertyInfo info, ParameterInfo? constructorParameter)
    {
        var syntax = info.Attribute<StringSyntaxAttribute>() ?? constructorParameter?.Attribute<StringSyntaxAttribute>();
        return syntax is not null &&
               string.Equals(syntax.Syntax, "html", StringComparison.OrdinalIgnoreCase);
    }

    static bool HasHtmlAttribute(PropertyInfo info, ParameterInfo? constructorParameter)
    {
        if (info.GetCustomAttributes(true).Any(IsHtmlAttribute))
        {
            return true;
        }

        return constructorParameter?.GetCustomAttributes(true).Any(IsHtmlAttribute) == true;

        static bool IsHtmlAttribute(object attribute) =>
            attribute.GetType().Name == "HtmlAttribute";
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

    static int? GetMinWidth(ColumnAttribute? column)
    {
        if (column is { MinWidth: > -1 })
        {
            return column.MinWidth;
        }

        return null;
    }

    static int? GetMaxWidth(ColumnAttribute? column)
    {
        if (column is { MaxWidth: > -1 })
        {
            return column.MaxWidth;
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
    public int? MinWidth { get; }
    public int? MaxWidth { get; }
    public string? Format { get; }
    public string? NullDisplay { get; }
    public bool IsHtml { get; }
    internal readonly bool IsHtmlExplicit;
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