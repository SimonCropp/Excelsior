class ColumnSettings
{
    Dictionary<string, ColumnSettings<IXLStyle>?> columns = [];

    public int? GetOrder(string name)
    {
        var config = columns.GetValueOrDefault(name);
        return config?.Order;
    }

    public void Add<T, TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<ColumnSettings<IXLStyle, TProperty>> configuration)
    {
        var name = property.PropertyName();
        var config = new ColumnSettings<IXLStyle, TProperty>();
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

    public bool TryGetValue(string name, [NotNullWhen(true)] out ColumnSettings<IXLStyle>? settings) =>
        columns.TryGetValue(name, out settings);
}