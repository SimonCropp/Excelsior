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
                Header = property.DisplayName,
                Width = null,
                HeaderStyle = null,
                CellStyle = null,
                Format = null,
                NullDisplay = null,
                Render = ValueRenderer.GetRender(property.Type),
                IsHtml = false,
                IsNumber = property.IsNumber,
                GetValue = _ => property.Get((TModel)_),
            };
        }
    }

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
        if (config.Header != null)
        {
            column.Header = config.Header;
        }

        if (config.Order != null)
        {
            column.Order = config.Order;
        }

        if (config.Width != null)
        {
            column.Width = config.Width;
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

        if (config.NullDisplay != null)
        {
            column.NullDisplay = config.NullDisplay;
        }

        if (config.Render != null)
        {
            column.Render = value => config.Render.Invoke((TProperty)value);
        }

        if (config.IsHtml != null)
        {
            column.IsHtml = config.IsHtml.Value;
        }
    }

    public List<Column<TStyle>> OrderedColumns() =>
        columns.Values.OrderBy(_ => _.Order ?? int.MaxValue).ToList();
}