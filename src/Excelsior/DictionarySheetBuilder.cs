namespace Excelsior;

class DictionarySheetBuilder :
    IDictionarySheetBuilder
{
    public List<ColumnConfig<IReadOnlyDictionary<string, object?>>> Columns { get; } = [];
    public bool AutoFilter { get; private set; } = true;

    public IDictionarySheetBuilder Column<TProperty>(
        string key,
        Action<DictionaryColumnConfig<TProperty>>? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            throw new ArgumentException("Column key must be supplied.", nameof(key));
        }

        var existing = Columns.FirstOrDefault(_ => _.Name == key);
        if (existing != null)
        {
            throw new($"Sheet already contains a column named '{key}'.");
        }

        var type = typeof(TProperty);
        var config = new DictionaryColumnConfig<TProperty>();
        configuration?.Invoke(config);

        var (isEnumerable, render) = ValueRenderer.GetRender(type);

        var allowedValues = config.DisableAllowedValues
            ? null
            : config.AllowedValues ?? TypeInference.DeriveAllowedValues(type);

        var keyForCapture = key;

        Func<IReadOnlyDictionary<string, object?>, object, string?>? finalRender;
        if (config.Render != null)
        {
            var userRender = config.Render;
            finalRender = (row, value) => userRender(row, (TProperty)value);
        }
        else if (!isEnumerable && render != null)
        {
            finalRender = (_, value) => render(value);
        }
        else
        {
            finalRender = null;
        }

        Action<CellStyle, IReadOnlyDictionary<string, object?>, object?>? finalCellStyle = null;
        if (config.CellStyle != null)
        {
            var userCellStyle = config.CellStyle;
            finalCellStyle = (style, row, value) => userCellStyle(style, row, (TProperty)value!);
        }

        var column = new ColumnConfig<IReadOnlyDictionary<string, object?>>
        {
            Name = key,
            Heading = config.Heading ?? key,
            Order = config.Order,
            DeclarationIndex = Columns.Count,
            Width = config.Width,
            MinWidth = config.MinWidth,
            MaxWidth = config.MaxWidth,
            HeadingStyle = config.HeadingStyle,
            CellStyle = finalCellStyle,
            Format = config.Format ?? DeriveDefaultFormat(type),
            NullDisplay = config.NullDisplay ?? ValueRenderer.GetNullDisplay(type),
            Render = finalRender,
            Formula = config.Formula,
            IsHtml = config.IsHtml ?? false,
            IsHtmlExplicit = config.IsHtml.HasValue,
            Filter = config.Filter,
            Include = config.Include ?? true,
            IsNumber = type.IsNumericType() ||
                       (Nullable.GetUnderlyingType(type)?.IsNumericType() ?? false),
            IsEnumerable = isEnumerable,
            ItemRender = isEnumerable ? render : null,
            GetValue = row => row.TryGetValue(keyForCapture, out var value) ? value : null,
            AllowedValues = allowedValues,
            NumericMin = config.NumericMin,
            NumericMax = config.NumericMax,
            DateMin = config.DateMin,
            DateMax = config.DateMax,
            Required = config.Required ?? false,
            Locked = config.Locked,
            InputTitle = config.InputTitle,
            InputMessage = config.InputMessage,
            ErrorTitle = config.ErrorTitle,
            ErrorMessage = config.ErrorMessage,
            ErrorStyle = config.ErrorStyle
        };

        Columns.Add(column);
        return this;
    }

    public void DisableFilter() => AutoFilter = false;

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

        if (underlying == typeof(Time))
        {
            return ValueRenderer.DefaultTimeFormat;
        }

        if (underlying == typeof(bool))
        {
            return ValueRenderer.BoolFormat;
        }

        return null;
    }

    internal List<ColumnConfig<IReadOnlyDictionary<string, object?>>> OrderedColumns() =>
        Columns
            .Where(_ => _.Include)
            .OrderBy(_ => _.Order.HasValue ? 0 : 1)
            .ThenBy(_ => _.Order ?? _.DeclarationIndex)
            .ToList();
}
