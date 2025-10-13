class Columns<TModel, TStyle>
{
    Dictionary<string, ColumnConfig<TStyle, TModel>> columns = [];

    public Columns()
    {
        foreach (var property in Properties<TModel>.Items)
        {
            AddColumn(property);
        }
    }

    void AddColumn(Property<TModel> property)
    {
        var type = property.Type;
        var render = ValueRenderer.GetRender(type);

        columns.Add(
            property.Name,
            new()
            {
                Name = property.Name,
                Order = property.Order,
                Heading = property.DisplayName,
                Width = property.Width,
                HeadingStyle = null,
                CellStyle = null,
                Format = property.Format,
                NullDisplay = property.NullDisplay ?? ValueRenderer.GetNullDisplay(type),
                Render = render == null ? null : (_, value) => render(value),
                IsHtml = property.IsHtml,
                IsNumber = property.IsNumber,
                IsEnumerableString = type.IsAssignableTo(typeof(IEnumerable<string>)),
                GetValue = property.Get
            });
    }

    public void Add<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnConfig<TStyle, TModel, TProperty>> configuration)
    {
        var name = property.PropertyName();
        if (!columns.TryGetValue(name, out var column))
        {
            throw new($"Could not find property: {name}");
        }

        var config = new ColumnConfig<TStyle, TModel, TProperty>();
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

    public List<ColumnConfig<TStyle, TModel>> OrderedColumns() =>
        columns.Values.OrderBy(_ => _.Order ?? int.MaxValue).ToList();
}