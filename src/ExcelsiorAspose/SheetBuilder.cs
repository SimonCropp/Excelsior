namespace ExcelsiorAspose;

public class SheetBuilder<T>(
    string name,
    IAsyncEnumerable<T> data,
    bool useAlternatingRowColors,
    Color? alternateRowColor,
    Action<Style>? headerStyle,
    Action<Style>? globalStyle,
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

        ApplyGlobalStyling(sheet);
        //TODO:
        //sheet.RangeUsed()!.SetAutoFilter();
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
            var cell = sheet.Cells[0, i];

            cell.Value = GetHeaderText(property);

            ApplyHeaderStyling(cell, property);
        }

        sheet.FreezePanes(1, 0, 1, 0);
    }

    async Task PopulateData(Sheet sheet, List<Property<T>> properties, Cancel cancel)
    {
        //Skip header
        var startRow = 1;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < properties.Count; colIndex++)
            {
                var property = properties[colIndex];
                var cell = sheet.Cells[xlRow, colIndex];

                var style = cell.GetStyle();
                style.VerticalAlignment = TextAlignmentType.Top;
                var value = property.Get(item);
                SetCellValue(cell, value, property, style);
                ApplyDataCellStyling(property, rowIndex, value, style);
                cell.SetStyle(style);
            }

            rowIndex++;
        }
    }

    void SetCellValue(Cell cell, object? value, Property<T> property, Style style)
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

            if (ValueRenderer.TryRender(property.Type, value, out var result))
            {
                cell.Value = result;
                return;
            }

            if (value is DateTime dateTime)
            {
                cell.Value = dateTime;
                if (config.Format != null)
                {
                    style.Custom = config.Format;
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
                cell.Value = enumValue.DisplayName();
                return;
            }

            if (property.IsNumber)
            {
                cell.Value = Convert.ToDouble(value);
                if (config.Format != null)
                {
                    style.Custom = config.Format;
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

            if (ValueRenderer.TryRender(property.Type, value, out var result))
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
                cell.Value = enumValue.DisplayName();
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
            WriteEnumerable(cell, enumerable, style);
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

    void WriteEnumerable(Cell cell, IEnumerable<string> enumerable, Style style)
    {
        style.IsTextWrapped = true;
        var list = enumerable.ToList();
        foreach (var item in list)
        {
            var builder = new StringBuilder("<ul>");
            foreach (var line in item.AsSpan().EnumerateLines())
            {
                builder.AppendLine("<li>");
                if (trimWhitespace)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    var encoded = WebUtility.HtmlEncode(line.Trim().ToString());
                    builder.Append(encoded);
                }
                else
                {
                    builder.Append(WebUtility.HtmlEncode(line.ToString()));
                }

                builder.AppendLine("</li>");
            }

            builder.AppendLine("</ul>");

            cell.HtmlString = builder.ToString();
        }
    }

    void ApplyHeaderStyling(Cell cell, Property<T> property)
    {
        var style = cell.GetStyle();
        // Apply global header styling
        headerStyle?.Invoke(style);

        if (settings.TryGetValue(property.Name, out var config))
        {
            config.HeaderStyle?.Invoke(style);
        }

        cell.SetStyle(style);
    }

    void ApplyDataCellStyling(Property<T> property, int index, object? value, Style style)
    {
        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.BackgroundColor = alternateRowColor!.Value;
        }

        if (!settings.TryGetValue(property.Name, out var config))
        {
            return;
        }

        config.DataCellStyle?.Invoke(style);
        config.ConditionalStyling?.Invoke(style, value);
    }

    void ApplyGlobalStyling(Sheet sheet)
    {
        if (globalStyle == null)
        {
            return;
        }

        var style = sheet.Workbook.CreateStyle();
        globalStyle(style);
        var flag = new StyleFlag
        {
            FontName = true,
            FontSize = true,
            FontColor = true,
            CellShading = true
        };
        sheet.Cells.ApplyStyle(style, flag);
    }

    void AutoSizeColumns(Sheet sheet, List<Property<T>> properties)
    {
        sheet.AutoFitColumns();

        //Round widths since aspose AutoFitColumns is not deterministic
        foreach (var column in sheet.Cells.Columns)
        {
            column.Width = Math.Round(column.Width);
        }

        // Apply specific column widths
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            if (settings.TryGetValue(property.Name, out var config) &&
                config.ColumnWidth.HasValue)
            {
                sheet.Cells.Columns[i].Width = config.ColumnWidth.Value;
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
}