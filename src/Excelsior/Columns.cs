class Columns<TModel, TStyle>
{
    Dictionary<string, Column<TStyle>> columns = [];

    public Columns()
    {
        foreach (var property in Properties<TModel>.Items)
        {
            columns[property.Name] = new()
            {
                Name = property.Name,
                Order = property.Order,
                HeaderText = property.DisplayName,
                ColumnWidth = null,
                HeaderStyle = null,
                CellStyle = null,
                Format = null,
                NullDisplayText = null,
                Render = ValueRenderer.GetRender(property.Type),
                TreatAsHtml = false,
                IsNumber = property.IsNumber,
                GetValue = _ => property.Get((TModel)_),
            };
        }
    }

    public List<Property<T>> ResolveProperties<T>() =>
        Properties<T>
            .Items
            .OrderBy(_ =>
            {
                if (columns.TryGetValue(_.Name, out var config) &&
                    config.Order != null)
                {
                    return config.Order.Value;
                }

                return _.Order ?? int.MaxValue;
            })
            .ToList();

    public void Add<T, TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<TStyle, TProperty>> configuration)
    {
        var name = property.PropertyName();
        if (!columns.TryGetValue(name, out var column))
        {
            throw new($"Could not find property: {name}");
        }

        var config = new Column<TStyle, TProperty>();
        configuration(config);
        if (config.HeaderText != null)
        {
            column.HeaderText = config.HeaderText;
        }

        if (config.Order != null)
        {
            column.Order = config.Order;
        }

        if (config.ColumnWidth != null)
        {
            column.ColumnWidth = config.ColumnWidth;
        }

        if (config.HeaderStyle != null)
        {
            column.HeaderStyle = config.HeaderStyle;
        }

        if (config.CellStyle != null)
        {
            column.CellStyle = (style, value) => config.CellStyle.Invoke(style, (TProperty)value!);
        }

        if (config.Format != null)
        {
            column.Format = config.Format;
        }

        if (config.NullDisplayText != null)
        {
            column.NullDisplayText = config.NullDisplayText;
        }

        if (config.Render != null)
        {
            column.Render = value => config.Render.Invoke((TProperty)value);
        }

        if (config.TreatAsHtml != null)
        {
            column.TreatAsHtml = config.TreatAsHtml.Value;
        }
    }

    public List<Column<TStyle>> OrderedColumns() =>
        columns.Values.OrderBy(_ => _.Order ?? int.MaxValue).ToList();
    public bool TryGetValue(string name, [NotNullWhen(true)] out Column<TStyle>? settings) =>
        columns.TryGetValue(name, out settings);

    public Column<TStyle> GetValue(string name) => columns[name];

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

    public bool TryGetHeaderStyle<T>(Property<T> property, [NotNullWhen(true)] out Action<TStyle>? headerStyle)
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