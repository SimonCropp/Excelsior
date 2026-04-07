class Columns<TModel>
{
    Dictionary<string, ColumnConfig<TModel>> columns = [];
    public bool AutoFilter { get; set; } = true;

    public Columns()
    {
        var index = 0;
        foreach (var property in Properties<TModel>.Items)
        {
            AddColumn(property, index++);
        }
    }

    void AddColumn(Property<TModel> property, int index)
    {
        var type = property.Type;
        var (isEnumerable, render) = ValueRenderer.GetRender(type);

        columns.Add(
            property.Name,
            new()
            {
                Name = property.Name,
                Order = property.Order,
                DeclarationIndex = index,
                Heading = property.DisplayName,
                Width = property.Width,
                HeadingStyle = null,
                CellStyle = null,
                Format = property.Format,
                NullDisplay = property.NullDisplay ?? ValueRenderer.GetNullDisplay(type),
                Render = isEnumerable ? null : render == null ? null : (_, value) => render(value),
                IsHtml = property.IsHtml,
                Filter = property.Filter,
                Include = property.Include ?? true,
                IsNumber = property.IsNumber,
                IsEnumerable = isEnumerable,
                ItemRender = isEnumerable ? render : null,
                GetValue = property.Get
            });
    }

    public void Add<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnConfig<TModel, TProperty>> configuration)
    {
        var name = property.PropertyName();
        if (!columns.TryGetValue(name, out var column))
        {
            throw new($"Could not find property: {name}");
        }

        var config = new ColumnConfig<TModel, TProperty>();
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

        if (config.Filter != null)
        {
            column.Filter = config.Filter.Value;
        }

        if (config.Include != null)
        {
            column.Include = config.Include.Value;
        }
    }

    public List<ColumnConfig<TModel>> OrderedColumns() =>
        columns.Values
            .Where(_ => _.Include)
            .OrderBy(_ => _.Order.HasValue ? 0 : 1)
            .ThenBy(_ => _.Order ?? _.DeclarationIndex)
            .ToList();
}
