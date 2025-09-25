namespace Excelsior;

public class SheetBuilder<T>(
    string name,
    IAsyncEnumerable<T> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<IXLStyle>? headerStyle,
    Action<IXLStyle>? globalStyle,
    bool trimWhitespace)
    where T : class
{
    static SheetBuilder() =>
        properties = typeof(T)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(_ => _.CanRead)
            .Select(_ => new Property<T>(_))
            .ToList();

    int rowIndex;
    Dictionary<string, ColumnSettings> settings = [];
    static IReadOnlyList<Property<T>> properties;

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<T> Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<ColumnSettings<TProperty>> configuration)
    {
        var name = property.PropertyName();
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
        };
        return this;
    }

    internal async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var properties = GetProperties();

        CreateHeaders(sheet, properties);

        await PopulateData(sheet, properties, cancel);

        ApplyGlobalStyling(sheet, properties);
        sheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(sheet, properties);
    }

    List<Property<T>> GetProperties() =>
        // Order by display order if specified
        properties
            .OrderBy(_ =>
            {
                var config = settings.GetValueOrDefault(_.Name);
                return config?.Order ?? _.Order ?? int.MaxValue;
            })
            .ToList();

    void CreateHeaders(Sheet sheet, List<Property<T>> properties)
    {
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = GetHeaderText(property);

            ApplyHeaderStyling(cell, property);
        }

        sheet.SheetView.FreezeRows(1);
    }

    async Task PopulateData(Sheet sheet, List<Property<T>> properties, Cancel cancel)
    {
        //Skip header
        var startRow = 2;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < properties.Count; colIndex++)
            {
                var property = properties[colIndex];
                var cell = sheet.Cell(xlRow, colIndex + 1);

                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                var value = property.Get(item);
                SetCellValue(cell, value, property);
                ApplyDataCellStyling(cell, property, rowIndex, value);
            }

            rowIndex++;
        }
    }

    void SetCellValue(Cell cell, object? value, Property<T> property)
    {
        if (settings.TryGetValue(property.Name, out var config))
        {
            if (value == null)
            {
                cell.Value = config.NullDisplayText;
                return;
            }

            if (config.Render != null)
            {
                cell.Value = config.Render(value);
                return;
            }

            if (BookBuilder.TryRender(property.Type, value, out var result))
            {
                cell.Value = result;
                return;
            }

            if (value is DateTime dateTime)
            {
                cell.Value = dateTime;
                if (config.Format != null)
                {
                    cell.Style.DateFormat.Format = config.Format;
                }

                return;
            }

            if (value is bool boolean)
            {
                cell.Value = boolean.ToString();
                return;
            }

            if (value is Enum enumValue)
            {
                cell.Value = GetEnumDisplayText(enumValue);
                return;
            }

            if (property.IsNumber)
            {
                cell.Value = Convert.ToDouble(value);
                if (config.Format != null)
                {
                    cell.Style.NumberFormat.Format = config.Format;
                }

                return;
            }
        }
        else
        {
            if (value == null)
            {
                cell.Value = "";
                return;
            }

            if (BookBuilder.TryRender(property.Type, value, out var result))
            {
                cell.Value = result;
                return;
            }

            if (value is DateTime dateTime)
            {
                cell.Value = dateTime;
                return;
            }

            if (value is bool boolean)
            {
                cell.Value = boolean.ToString();
                return;
            }

            if (value is Enum enumValue)
            {
                cell.Value = GetEnumDisplayText(enumValue);
                return;
            }

            if (property.IsNumber)
            {
                cell.Value = Convert.ToDouble(value);
                return;
            }
        }

        if (value is IEnumerable<string> enumerable)
        {
            WriteEnumerable(cell, enumerable);
            return;
        }

        var valueAsString = value.ToString();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
        if (valueAsString != null && trimWhitespace)
        {
            valueAsString = valueAsString.Trim();
        }
        cell.Value = valueAsString;
    }

    void WriteEnumerable(Cell cell, IEnumerable<string> enumerable)
    {
        cell.Style.Alignment.WrapText = true;
        var rich = cell.CreateRichText();
        var list = enumerable.ToList();
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            rich.AddText("• ").SetBold();
            var builder = new StringBuilder();
            foreach (var line in item.AsSpan().EnumerateLines())
            {
                if (trimWhitespace)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }
                    builder.Append(line.Trim());
                }
                else
                {
                    builder.Append(line);
                }

                builder.Append("\n   ");
            }

            if (index < list.Count - 1)
            {
                builder.Length -= 3;
            }
            else
            {
                builder.Length -= 4;
            }

            rich.AddText(builder.ToString());
        }
    }

    void ApplyHeaderStyling(Cell cell, Property<T> property)
    {
        // Apply global header styling
        headerStyle?.Invoke(cell.Style);

        if (settings.TryGetValue(property.Name, out var config))
        {
            config.HeaderStyle?.Invoke(cell.Style);
        }
    }

    void ApplyDataCellStyling(Cell cell, Property<T> property, int index, object? value)
    {
        var style = cell.Style;

        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.Fill.BackgroundColor = alternateRowColor;
        }

        if (!settings.TryGetValue(property.Name, out var config))
        {
            return;
        }

        config.DataCellStyle?.Invoke(style);
        config.ConditionalStyling?.Invoke(style, value);
    }

    void ApplyGlobalStyling(Sheet sheet, List<Property<T>> properties)
    {
        if (globalStyle == null)
        {
            return;
        }

        var range = sheet.Range(1, 1, rowIndex + 1, properties.Count);
        globalStyle(range.Style);
    }

    void AutoSizeColumns(Sheet sheet, List<Property<T>> properties)
    {
        sheet.Columns().AdjustToContents();

        // Apply specific column widths
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            if (settings.TryGetValue(property.Name, out var config) &&
                config.ColumnWidth.HasValue)
            {
                sheet.Column(i + 1).Width = config.ColumnWidth.Value;
            }
        }
    }

    string GetHeaderText(Property<T> property)
    {
        if (settings.TryGetValue(property.Name, out var config) &&
            config.HeaderText != null)
        {
            return config.HeaderText;
        }

        return property.DisplayName;
    }

    static string GetEnumDisplayText(Enum enumValue)
    {
        var field = enumValue.GetType().GetField(enumValue.ToString());
        var attribute = field?.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Name ?? enumValue.ToString();
    }
}