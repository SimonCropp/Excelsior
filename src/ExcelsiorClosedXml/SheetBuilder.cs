namespace ExcelsiorClosedXml;

public class SheetBuilder<T>(
    string name,
    IAsyncEnumerable<T> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<IXLStyle>? headerStyle,
    Action<IXLStyle>? globalStyle,
    bool trimWhitespace) :
    ISheetBuilder<T, IXLStyle>
    where T : class
{
    int rowIndex;
    Columns<IXLStyle> columns = new();

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<T> Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<IXLStyle, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    void ISheetBuilder<T, IXLStyle>.Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<IXLStyle, TProperty>> configuration) =>
        Column(property, configuration);

    internal async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var properties = columns.ResolveProperties<T>();

        CreateHeaders(sheet, properties);

        await PopulateData(sheet, properties, cancel);

        ApplyGlobalStyling(sheet, properties);
        sheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(sheet, properties);
    }

    void CreateHeaders(Sheet sheet, List<Property<T>> properties)
    {
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = columns.GetHeaderText(property);

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
                cell.Value = enumValue.DisplayName();
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

        if (columns.TryGetHeaderStyle(property, out var columnHeaderStyle))
        {
            columnHeaderStyle.Invoke(cell.Style);
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

        if (!columns.TryGetValue(property.Name, out var config))
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
            if (columns.TryGetColumnWidth(properties[i], out var width))
            {
                sheet.Column(i + 1).Width = width;
            }
        }
    }
}