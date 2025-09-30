class Columns<TModel, TStyle>
{
    Dictionary<string, Column<TStyle, TModel>> columns = [];

    public Columns()
    {
        foreach (var property in Properties<TModel>.Items)
        {
            var render = ValueRenderer.GetRender(property.Type);
            Func<TModel, object, string?>? renderFunc;
            if (render == null)
            {
                renderFunc = null;
            }
            else
            {
                renderFunc = (_, value) => render(value);
            }

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
                Render = renderFunc,
                IsHtml = false,
                IsNumber = property.IsNumber,
                GetValue = _ => property.Get(_),
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