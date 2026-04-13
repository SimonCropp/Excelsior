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
                MinWidth = property.MinWidth,
                MaxWidth = property.MaxWidth,
                HeadingStyle = null,
                CellStyle = null,
                Format = property.Format,
                NullDisplay = property.NullDisplay ?? ValueRenderer.GetNullDisplay(type),
                Render = isEnumerable ? null : render == null ? null : (_, value) => render(value),
                Formula = null,
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

        if (config.MinWidth != null)
        {
            column.MinWidth = config.MinWidth;
        }

        if (config.MaxWidth != null)
        {
            column.MaxWidth = config.MaxWidth;
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

        if (config.Formula != null)
        {
            column.Formula = config.Formula;
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

    const int MaxExcelColumnWidth = 255;

    public List<ColumnConfig<TModel>> OrderedColumns()
    {
        foreach (var column in columns.Values)
        {
            ValidateWidth(column, column.Width, "Width");
            ValidateWidth(column, column.MinWidth, "MinWidth");
            ValidateWidth(column, column.MaxWidth, "MaxWidth");

            if (column.Width is not null && (column.MinWidth is not null || column.MaxWidth is not null))
            {
                throw new($"Column '{column.Name}': Width cannot be combined with MinWidth/MaxWidth. Use either Width for an exact size, or MinWidth/MaxWidth for auto-sizing with bounds.");
            }

            if (column is not { MinWidth: { } min, MaxWidth: { } max })
            {
                continue;
            }

            if (min == max)
            {
                throw new($"Column '{column.Name}': MinWidth and MaxWidth are both {min}. Use Width instead.");
            }

            if (min > max)
            {
                throw new($"Column '{column.Name}': MinWidth ({min}) is greater than MaxWidth ({max}).");
            }
        }

        return columns.Values
            .Where(_ => _.Include)
            .OrderBy(_ => _.Order.HasValue ? 0 : 1)
            .ThenBy(_ => _.Order ?? _.DeclarationIndex)
            .ToList();
    }

    static void ValidateWidth(ColumnConfig<TModel> column, int? value, string name)
    {
        if (value > MaxExcelColumnWidth)
        {
            throw new($"Column '{column.Name}': {name} ({value}) exceeds the Excel maximum of {MaxExcelColumnWidth}.");
        }
    }
}
