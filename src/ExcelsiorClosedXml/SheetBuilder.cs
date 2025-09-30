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
    Columns<T, IXLStyle> columns = new();

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

        var orderedColumns = columns.OrderedColumns();
        CreateHeaders(sheet, orderedColumns);

        await PopulateData(sheet, orderedColumns, cancel);

        ApplyGlobalStyling(sheet, orderedColumns);
        sheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(sheet, orderedColumns);
    }

    void CreateHeaders(Sheet sheet, List<Column<IXLStyle>> orderedColumns)
    {
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = column.HeaderText;

            ApplyHeaderStyling(cell, column);
        }

        sheet.SheetView.FreezeRows(1);
    }

    async Task PopulateData(Sheet sheet, List<Column<IXLStyle>> orderedColumns, Cancel cancel)
    {
        //Skip header
        var startRow = 2;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < orderedColumns.Count; colIndex++)
            {
                var column = orderedColumns[colIndex];
                var cell = sheet.Cell(xlRow, colIndex + 1);

                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cell.Style.Alignment.WrapText = true;

                var value = column.GetValue(item);
                SetCellValue(cell, value, column);
                ApplyCellStyle(cell, rowIndex, value, column);
            }

            rowIndex++;
        }
    }

    void SetCellValue(Cell cell, object? value, Column<IXLStyle> column)
    {
        if (value == null)
        {
            cell.Value = column.NullDisplayText;
            return;
        }

        if (column.Render != null)
        {
            cell.Value = column.Render(value);
            return;
        }

        if (value is DateTime dateTime)
        {
            cell.Value = dateTime;
            if (column.Format != null)
            {
                cell.Style.DateFormat.Format = column.Format;
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

        if (column.IsNumber)
        {
            cell.Value = Convert.ToDouble(value);
            if (column.Format != null)
            {
                cell.Style.NumberFormat.Format = column.Format;
            }

            return;
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

    void ApplyHeaderStyling(Cell cell, Column<IXLStyle> column)
    {
        headerStyle?.Invoke(cell.Style);

        column.HeaderStyle?.Invoke(cell.Style);
    }

    void ApplyCellStyle(Cell cell, int index, object? value, Column<IXLStyle> column)
    {
        var style = cell.Style;

        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.Fill.BackgroundColor = alternateRowColor;
        }

        column.CellStyle?.Invoke(style, value);
    }

    void ApplyGlobalStyling(Sheet sheet, List<Column<IXLStyle>> orderedColumns)
    {
        if (globalStyle == null)
        {
            return;
        }

        var range = sheet.Range(1, 1, rowIndex + 1, orderedColumns.Count);
        globalStyle(range.Style);
    }

    static void AutoSizeColumns(Sheet sheet, List<Column<IXLStyle>> orderedColumns)
    {
        var xlColumns = sheet.Columns().ToList();

        // Apply specific column widths
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            if (column.Width != null)
            {
                var xlColumn = sheet.Column(i + 1);
                xlColumns.Remove(xlColumn);
                xlColumn.Width = column.Width.Value;
            }
        }

        var endRow = sheet.RowsUsed().Count();
        foreach (var column in xlColumns)
        {
            column.AdjustToContents(1, endRow);
            column.Width += 2;
        }
    }
}