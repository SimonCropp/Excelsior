class Columns<TStyle>
{
    Dictionary<string, Column<TStyle>> columns = [];

    public List<Property<T>> ResolveProperties<T>() =>
        Properties<T>
            .Items
            .OrderBy(_ =>
            {
                if (columns.TryGetValue(_.Name, out var config))
                {
                    if (config.Order != null)
                    {
                        return config.Order.Value;
                    }
                }

                return _.Order ?? int.MaxValue;
            })
            .ToList();

    public void Add<T, TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<TStyle, TProperty>> configuration)
    {
        var name = property.PropertyName();
        var config = new Column<TStyle, TProperty>();
        configuration(config);
        Func<object, string?>? render;
        if (config.Render == null)
        {
            render = null;
        }
        else
        {
            render = o => config.Render.Invoke((TProperty) o);
        }

        columns[name] = new()
        {
            HeaderText = config.HeaderText,
            Order = config.Order,
            ColumnWidth = config.ColumnWidth,
            HeaderStyle = config.HeaderStyle,
            DataCellStyle = config.DataCellStyle,
            ConditionalStyling = (style, o) => config.ConditionalStyling?.Invoke(style, (TProperty) o!),
            Format = config.Format,
            NullDisplayText = config.NullDisplayText,
            Render = render,
        };
    }

    public bool TryGetValue(string name, [NotNullWhen(true)] out Column<TStyle>? settings) =>
        columns.TryGetValue(name, out settings);

    public string GetHeaderText<T>(Property<T> property)
    {
        if (TryGetValue(property.Name, out var config) &&
            config.HeaderText != null)
        {
            return config.HeaderText;
        }

        return property.DisplayName;
    }

    public bool TryGetColumnWidth<T>(Property<T> property, out double width)
    {
        if (TryGetValue(property.Name, out var config) &&
            config.ColumnWidth.HasValue)
        {
            width = config.ColumnWidth.Value;
            return true;
        }

        width = 0;
        return false;
    }

    public bool TryGetHeaderStyle<T>(Property<T> property, [NotNullWhen(true)]out Action<TStyle>? headerStyle)
    {
        if (TryGetValue(property.Name, out var config) &&
            config.HeaderStyle != null)
        {
            headerStyle = config.HeaderStyle;
            return true;
        }

        headerStyle = null;
        return false;
    }
}