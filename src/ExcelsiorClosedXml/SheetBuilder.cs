namespace ExcelsiorClosedXml;

public class SheetBuilder<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<Style>? headerStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace) :
    ISheetBuilder<TModel, Style, Cell>
    where TModel : class
{
    int rowIndex;
    Columns<TModel, Style> columns = new();

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<TModel> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<Style, TModel, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    void ISheetBuilder<TModel, Style, Cell>.Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<Style, TModel, TProperty>> configuration) =>
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

    void CreateHeaders(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
    {
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = column.Header;

            ApplyHeaderStyling(cell, column);
        }

        sheet.SheetView.FreezeRows(1);
    }

    async Task PopulateData(Sheet sheet, List<Column<Style, TModel>> orderedColumns, Cancel cancel)
    {
        //Skip header
        var startRow = 2;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var index = 0; index < orderedColumns.Count; index++)
            {
                var column = orderedColumns[index];
                var cell = sheet.Cell(xlRow, index + 1);

                var style = cell.Style;
                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                style.Alignment.WrapText = true;

                var value = column.GetValue(item);
                ((ISheetBuilder<TModel, Style, Cell>) this).SetCellValue(cell, style, value, column, item, trimWhitespace);
                ApplyCellStyle(cell, rowIndex, value, column, item);
            }

            rowIndex++;
        }
    }

    void ISheetBuilder<TModel, Style, Cell>.SetDateFormat(Style style, string format) =>
        style.DateFormat.Format = format;

    void ISheetBuilder<TModel, Style, Cell>.SetNumberFormat(Style style, string format) =>
        style.NumberFormat.Format = format;

    void ISheetBuilder<TModel, Style, Cell>.SetCellValue(Cell cell, object value) =>
        cell.Value = XLCellValue.FromObject(value);

    void ISheetBuilder<TModel, Style, Cell>.SetCellHtml(Cell cell, string value) =>
        throw new ("ClosedXml does not support html");

    void ISheetBuilder<TModel, Style, Cell>.WriteEnumerable(Cell cell, IEnumerable<string> enumerable)
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

    void ApplyHeaderStyling(Cell cell, Column<Style, TModel> column)
    {
        headerStyle?.Invoke(cell.Style);

        column.HeaderStyle?.Invoke(cell.Style);
    }

    void ApplyCellStyle(Cell cell, int index, object? value, Column<Style, TModel> column, TModel item)
    {
        var style = cell.Style;

        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.Fill.BackgroundColor = alternateRowColor;
        }

        column.CellStyle?.Invoke(style, item, value);
    }

    void ApplyGlobalStyling(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
    {
        if (globalStyle == null)
        {
            return;
        }

        var range = sheet.Range(1, 1, rowIndex + 1, orderedColumns.Count);
        globalStyle(range.Style);
    }

    static void AutoSizeColumns(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
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