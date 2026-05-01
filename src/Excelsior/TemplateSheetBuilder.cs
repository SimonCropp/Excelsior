namespace Excelsior;

class TemplateSheetBuilder :
    ITemplateSheetBuilder
{
    public List<ColumnConfig<TemplateRow>> Columns { get; } = [];
    public bool AutoFilter { get; private set; } = true;

    public ITemplateSheetBuilder Column<TProperty>(
        string name,
        Action<TemplateColumnConfig<TProperty>>? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Column name must be supplied.", nameof(name));
        }

        var existing = Columns.FirstOrDefault(_ => _.Name == name);
        if (existing != null)
        {
            throw new($"Template sheet already contains a column named '{name}'.");
        }

        var type = typeof(TProperty);
        var config = new TemplateColumnConfig<TProperty>();
        configuration?.Invoke(config);

        var allowedValues = config.DisableAllowedValues
            ? null
            : config.AllowedValues ?? DeriveEnumAllowedValues(type);

        var column = new ColumnConfig<TemplateRow>
        {
            Name = name,
            Heading = config.Heading ?? name,
            Order = config.Order,
            DeclarationIndex = Columns.Count,
            Width = config.Width,
            MinWidth = config.MinWidth,
            MaxWidth = config.MaxWidth,
            HeadingStyle = config.HeadingStyle,
            CellStyle = null,
            Format = config.Format ?? DeriveDefaultFormat(type),
            NullDisplay = null,
            Render = null,
            Formula = null,
            IsHtml = false,
            IsHtmlExplicit = false,
            Filter = config.Filter,
            Include = true,
            IsNumber = type.IsNumericType() ||
                       (Nullable.GetUnderlyingType(type)?.IsNumericType() ?? false),
            IsEnumerable = false,
            ItemRender = null,
            GetValue = _ => null,
            AllowedValues = allowedValues,
            NumericMin = config.NumericMin,
            NumericMax = config.NumericMax,
            DateMin = config.DateMin,
            DateMax = config.DateMax,
            Required = config.Required,
            Locked = config.Locked,
            InputTitle = config.InputTitle,
            InputMessage = config.InputMessage,
            ErrorTitle = config.ErrorTitle,
            ErrorMessage = config.ErrorMessage
        };
        Columns.Add(column);
        return this;
    }

    public void DisableFilter() => AutoFilter = false;

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

    static string? DeriveDefaultFormat(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type) ?? type;
        if (underlying == typeof(DateTime))
        {
            return ValueRenderer.DefaultDateTimeFormat;
        }

        if (underlying == typeof(DateTimeOffset))
        {
            return ValueRenderer.DefaultDateTimeOffsetFormat;
        }

        if (underlying == typeof(Date))
        {
            return ValueRenderer.DefaultDateFormat;
        }

        return null;
    }

    internal List<ColumnConfig<TemplateRow>> OrderedColumns() =>
        Columns
            .OrderBy(_ => _.Order.HasValue ? 0 : 1)
            .ThenBy(_ => _.Order ?? _.DeclarationIndex)
            .ToList();
}
