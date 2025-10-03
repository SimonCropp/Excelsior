namespace ExcelsiorClosedXml;

public class SheetBuilder<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<Style>? headingStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace) :
    SheetBuilderBase<TModel, Style, Cell>,
    ISheetBuilder<TModel, Style>
{
    int rowIndex;
    Columns<TModel, Style> columns = new();

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public override ISheetBuilder<TModel, Style> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<Style, TModel, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    internal async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var orderedColumns = columns.OrderedColumns();
        CreateHeadings(sheet, orderedColumns);

        await PopulateData(sheet, orderedColumns, cancel);

        ApplyGlobalStyling(sheet, orderedColumns);
        sheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(sheet, orderedColumns);
    }

    void CreateHeadings(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
    {
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = column.Heading;

            ApplyHeadingStyling(cell, column);
        }

        sheet.SheetView.FreezeRows(1);
    }

    async Task PopulateData(Sheet sheet, List<Column<Style, TModel>> orderedColumns, Cancel cancel)
    {
        //Skip heading
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
                base.SetCellValue(cell, style, value, column, item, trimWhitespace);
                ApplyCellStyle(cell, rowIndex, value, column, item);
            }

            rowIndex++;
        }
    }

    protected override void SetDateFormat(Style style, string format) =>
        style.DateFormat.Format = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat.Format = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value = XLCellValue.FromObject(value);

    protected override void SetCellHtml(Cell cell, string value) =>
        throw new ("ClosedXml does not support html");

    protected override void WriteEnumerable(Cell cell, IEnumerable<string> enumerable)
    {
        var rich = cell.CreateRichText();
        var list = enumerable.ToList();
        var builder = new StringBuilder();
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            rich.AddText("• ").SetBold();
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
            builder.Clear();
        }
    }

    void ApplyHeadingStyling(Cell cell, Column<Style, TModel> column)
    {
        headingStyle?.Invoke(cell.Style);

        column.HeadingStyle?.Invoke(cell.Style);
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