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

        var allowedValues = DeriveEnumAllowedValues(type);

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
                IsHtmlExplicit = property.IsHtmlExplicit,
                Filter = property.Filter,
                Include = property.Include ?? true,
                IsNumber = property.IsNumber,
                IsEnumerable = isEnumerable,
                ItemRender = isEnumerable ? render : null,
                GetValue = property.Get,
                AllowedValues = allowedValues
            });
    }

    static IReadOnlyList<string>? DeriveEnumAllowedValues(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type) ?? type;
        if (!underlying.IsEnum)
        {
            return null;
        }

        var (_, render) = ValueRenderer.GetRender(underlying);
        var values = Enum.GetValues(underlying);
        var list = new List<string>(values.Length);
        foreach (var value in values)
        {
            list.Add(render?.Invoke(value!) ?? value!.ToString()!);
        }

        return list;
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
            // A custom render produces cell values that won't match the auto-derived enum
            // dropdown list, so suppress it unless the caller has supplied an explicit list.
            if (config.AllowedValues == null && !config.DisableAllowedValues)
            {
                column.AllowedValues = null;
            }
        }

        if (config.Formula != null)
        {
            column.Formula = config.Formula;
            if (config.AllowedValues == null && !config.DisableAllowedValues)
            {
                column.AllowedValues = null;
            }
        }

        if (config.IsHtml is { } fluentIsHtml)
        {
            if (column.IsHtmlExplicit && column.IsHtml != fluentIsHtml)
            {
                throw new($"Column '{column.Name}': mismatched IsHtml — attribute says {column.IsHtml}, fluent configuration says {fluentIsHtml}.");
            }

            column.IsHtml = fluentIsHtml;
            column.IsHtmlExplicit = true;
        }

        if (config.Filter != null)
        {
            column.Filter = config.Filter.Value;
        }

        if (config.Include != null)
        {
            column.Include = config.Include.Value;
        }

        if (config.DisableAllowedValues)
        {
            column.AllowedValues = null;
        }
        else if (config.AllowedValues != null)
        {
            column.AllowedValues = config.AllowedValues;
        }

        if (config.NumericMin.HasValue)
        {
            column.NumericMin = config.NumericMin;
        }

        if (config.NumericMax.HasValue)
        {
            column.NumericMax = config.NumericMax;
        }

        if (config.DateMin.HasValue)
        {
            column.DateMin = config.DateMin;
        }

        if (config.DateMax.HasValue)
        {
            column.DateMax = config.DateMax;
        }

        if (config.Required)
        {
            column.Required = true;
        }

        if (config.Locked.HasValue)
        {
            column.Locked = config.Locked;
        }

        if (config.InputTitle != null)
        {
            column.InputTitle = config.InputTitle;
        }

        if (config.InputMessage != null)
        {
            column.InputMessage = config.InputMessage;
        }

        if (config.ErrorTitle != null)
        {
            column.ErrorTitle = config.ErrorTitle;
        }

        if (config.ErrorMessage != null)
        {
            column.ErrorMessage = config.ErrorMessage;
        }
    }

    const int maxExcelColumnWidth = 255;

    public List<ColumnConfig<TModel>> OrderedColumns()
    {
        foreach (var column in columns.Values)
        {
            ValidateWidth(column, column.Width, "Width");
            ValidateWidth(column, column.MinWidth, "MinWidth");
            ValidateWidth(column, column.MaxWidth, "MaxWidth");

            if (column.Width is not null &&
                (column.MinWidth is not null ||
                 column.MaxWidth is not null))
            {
                throw new($"Column '{column.Name}': Width cannot be combined with MinWidth/MaxWidth. Use either Width for an exact size, or MinWidth/MaxWidth for auto-sizing with bounds.");
            }

            if (column is not
                {
                    MinWidth: { } min,
                    MaxWidth: { } max
                })
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

        return columns
            .Values
            .Where(_ => _.Include)
            .OrderBy(_ => _.Order.HasValue ? 0 : 1)
            .ThenBy(_ => _.Order ?? _.DeclarationIndex)
            .ToList();
    }

    static void ValidateWidth(ColumnConfig<TModel> column, int? value, string name)
    {
        if (value > maxExcelColumnWidth)
        {
            throw new($"Column '{column.Name}': {name} ({value}) exceeds the Excel maximum of {maxExcelColumnWidth}.");
        }
    }
}
