class Columns<TStyle>
{
    Dictionary<string, Column<TStyle>?> columns = [];

    public int? GetOrder(string name)
    {
        var config = columns.GetValueOrDefault(name);
        return config?.Order;
    }

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
}