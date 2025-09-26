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
    int rowIndex;
    Columns<Style> columns = new();

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<T> Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<Style, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    internal async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var properties = columns.ResolveProperties<T>();

        CreateHeaders(sheet, properties);

        await PopulateData(sheet, properties, cancel);

        ApplyGlobalStyling(sheet);
        sheet.AutoFilterAll();
        AutoSizeColumns(sheet, properties);
        sheet.AutoSizeRows();
    }

    void CreateHeaders(Sheet sheet, List<Property<T>> properties)
    {
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var cell = sheet.Cells[0, i];

            cell.Value = columns.GetHeaderText(property);

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
        if (columns.TryGetValue(property.Name, out var config))
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
        var builder = new StringBuilder("<ul>");
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            builder.AppendLine("<li>");

            if (trimWhitespace)
            {
                item = item.Trim();
            }

            builder.Append(WebUtility.HtmlEncode(item));

            builder.AppendLine("</li>");
        }

        builder.AppendLine("</ul>");

        cell.HtmlString = builder.ToString();
    }

    void ApplyHeaderStyling(Cell cell, Property<T> property)
    {
        var style = cell.GetStyle();
        // Apply global header styling
        headerStyle?.Invoke(style);

        if (columns.TryGetHeaderStyle(property, out var columnHeaderStyle))
        {
            columnHeaderStyle.Invoke(style);
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

        if (!columns.TryGetValue(property.Name, out var config))
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
        sheet.AutoSizeColumns();

        // Apply specific column widths
        for (var i = 0; i < properties.Count; i++)
        {
            if (columns.TryGetColumnWidth(properties[i], out var width))
            {
                sheet.Cells.Columns[i].Width = width;
            }
        }
    }
}