namespace Excelsior;

public class SheetBuilder<T>(
    string name,
    IAsyncEnumerable<T> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<IXLStyle>? headerStyle,
    Action<IXLStyle>? globalStyle)
    where T : class
{
    static SheetBuilder() =>
        properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(_ => _.CanRead)
            .Select(_=>new Property(_))
            .ToList();

    int rowIndex;
    Dictionary<string, ColumnSettings> settings = [];
    static IReadOnlyList<Property> properties;

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<T> Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<ColumnSettings<TProperty>> configuration)
    {
        var name = GetPropertyName(property);
        var config = new ColumnSettings<TProperty>();
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

        settings[name] = new()
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
        } ;
        return this;
    }

    internal async Task AddSheet(XLWorkbook workbook, Cancel cancel)
    {
        var worksheet = workbook.Worksheets.Add(name);

        var properties = GetProperties();

        CreateHeaders(worksheet, properties);

        await PopulateData(worksheet, properties, cancel);

        ApplyGlobalStyling(worksheet, properties);
        worksheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(worksheet, properties);
    }

    List<PropertyInfo> GetProperties() =>
        // Order by display order if specified
        properties
            .OrderBy(_ =>
            {
                var config = settings.GetValueOrDefault(_.Info.Name);
                return config?.Order ?? GetDisplayOrder(_.Info) ?? int.MaxValue;
            })
            .Select(_=>_.Info)
            .ToList();

    void CreateHeaders(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var cell = worksheet.Cell(1, i + 1);

            cell.Value = GetHeaderText(property);

            ApplyHeaderStyling(cell, property);
        }

        worksheet.SheetView.FreezeRows(1);
    }

    async Task PopulateData(IXLWorksheet worksheet, List<PropertyInfo> properties, Cancel cancel)
    {
        //Skip header
        var startRow = 2;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < properties.Count; colIndex++)
            {
                var property = properties[colIndex];
                var cell = worksheet.Cell(xlRow, colIndex + 1);

                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                var value = property.GetValue(item);
                SetCellValue(cell, value, property);
                ApplyDataCellStyling(cell, property, rowIndex, value);
            }

            rowIndex++;
        }
    }

    void SetCellValue(IXLCell cell, object? value, PropertyInfo property)
    {
        if (value == null)
        {
            var config = settings.GetValueOrDefault(property.Name);
            cell.Value = config?.NullDisplayText ?? "";
        }
        else
        {
            var config = settings.GetValueOrDefault(property.Name);

            // Apply custom formatter if provided
            if (config?.Render != null)
            {
                cell.Value = config.Render(value);
                return;
            }

            if (BookBuilder.TryRender(property.PropertyType, value, out var result))
            {
                cell.Value = result;
                return;
            }

            void AssignNumberFormat()
            {
                if (config?.Format != null)
                {
                    cell.Style.NumberFormat.Format = config.Format;
                }
            }

            switch (value)
            {
                case DateTime dateTime:
                    cell.Value = dateTime;
                    if (config?.Format != null)
                    {
                        cell.Style.DateFormat.Format = config.Format;
                    }
                    break;

                case decimal decimalValue:
                    cell.Value = decimalValue;
                    AssignNumberFormat();
                    break;
                case double doubleValue:
                    cell.Value = doubleValue;
                    AssignNumberFormat();
                    break;
                case float floatValue:
                    cell.Value = floatValue;
                    AssignNumberFormat();
                    break;

                case bool boolean:
                    cell.Value = config?.Render?.Invoke(boolean) ?? boolean.ToString();
                    break;

                case Enum enumValue:
                    cell.Value = config?.Render?.Invoke(enumValue) ?? GetEnumDisplayText(enumValue);
                    break;

                default:
                    cell.Value = value.ToString();
                    break;
            }
        }
    }

    void ApplyHeaderStyling(IXLCell cell, PropertyInfo property)
    {
        var config = settings.GetValueOrDefault(property.Name);

        // Apply global header styling
        headerStyle?.Invoke(cell.Style);

        // Apply column-specific header styling
        config?.HeaderStyle?.Invoke(cell.Style);
    }

    void ApplyDataCellStyling(IXLCell cell, PropertyInfo property, int rowIndex, object? value)
    {
        var style = cell.Style;

        // Apply alternating row colors
        if (useAlternatingRowColors && rowIndex % 2 == 1)
        {
            style.Fill.BackgroundColor = alternateRowColor;
        }

        var config = settings.GetValueOrDefault(property.Name);

        if (config == null)
        {
            return;
        }

        config.DataCellStyle?.Invoke(style);
        config.ConditionalStyling?.Invoke(style, value);
    }

    void ApplyGlobalStyling(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        if (globalStyle == null)
        {
            return;
        }

        var range = worksheet.Range(1, 1, rowIndex + 1, properties.Count);
        globalStyle(range.Style);
    }

    void AutoSizeColumns(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        worksheet.Columns().AdjustToContents();

        // Apply specific column widths
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var config = settings.GetValueOrDefault(property.Name);

            if (config?.ColumnWidth.HasValue == true)
            {
                worksheet.Column(i + 1).Width = config.ColumnWidth.Value;
            }
        }
    }

    string GetHeaderText(PropertyInfo property)
    {
        var config = settings.GetValueOrDefault(property.Name);
        if (config?.HeaderText != null)
        {
            return config.HeaderText;
        }

        // Check for DisplayAttribute
        var display = property.GetCustomAttribute<DisplayAttribute>();
        if (display?.Name != null)
        {
            return display.Name;
        }

        // Check for DisplayNameAttribute
        var displayName = property.GetCustomAttribute<DisplayNameAttribute>();
        if (displayName != null)
        {
            return displayName.DisplayName;
        }

        return CamelCase.Split(property.Name);
    }

    static int? GetDisplayOrder(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Order;
    }

    static string GetEnumDisplayText(Enum enumValue)
    {
        var field = enumValue.GetType().GetField(enumValue.ToString());
        var attribute = field?.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Name ?? enumValue.ToString();
    }

    static string GetPropertyName<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
        if (propertyExpression.Body is MemberExpression member)
        {
            return member.Member.Name;
        }

        throw new ArgumentException("Expression must be a property access", nameof(propertyExpression));
    }
}