class Columns<TModel, TStyle>
{
    Dictionary<string, Column<TStyle, TModel>> columns = [];

    public Columns()
    {
        foreach (var property in Properties<TModel>.Items)
        {
            var render = ValueRenderer.GetRender(property.Type);

            columns[property.Name] = new()
            {
                Name = property.Name,
                Order = property.Order,
                Heading = property.DisplayName,
                Width = property.Width,
                HeadingStyle = null,
                CellStyle = null,
                Format = property.Format,
                NullDisplay = property.NullDisplay,
                Render = render == null ? null : (_, value) => render(value),
                IsHtml = property.IsHtml,
                IsNumber = property.IsNumber,
                GetValue = property.Get
            };
        }
    }

    public void Add<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<TStyle, TModel, TProperty>> configuration)
    {
        var name = property.PropertyName();
        if (!columns.TryGetValue(name, out var column))
        {
            throw new($"Could not find property: {name}");
        }

        var config = new Column<TStyle, TModel, TProperty>();
        configuration(config);
        if (config.Heading != null)
        {
            column.Heading = config.Heading;
        }

        if (config.Order != null)
        {
            column.Order = config.Order;
        }

        if (config.Width != null)
        {
            column.Width = config.Width;
        }

        if (config.HeadingStyle != null)
        {
            column.HeadingStyle = config.HeadingStyle;
        }

        if (config.CellStyle != null)
        {
            column.CellStyle = (style, model, value) => config.CellStyle.Invoke(style, model, (TProperty)value!);
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
            column.Render = (model, value) => config.Render.Invoke(model, (TProperty)value);
        }

        if (config.IsHtml != null)
        {
            column.IsHtml = config.IsHtml.Value;
        }
    }

    public List<Column<TStyle, TModel>> OrderedColumns() =>
        columns.Values.OrderBy(_ => _.Order ?? int.MaxValue).ToList();
}