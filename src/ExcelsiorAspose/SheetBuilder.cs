namespace ExcelsiorAspose;

public class SheetBuilder<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    Color? alternateRowColor,
    Action<Style>? headingStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace) :
    SheetBuilderBase<TModel, Style, Cell, Book>
{
    int rowIndex;

    internal override async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var orderedColumns = Columns.OrderedColumns();
        CreateHeadings(sheet, orderedColumns);

        await PopulateData(sheet, orderedColumns, cancel);

        ApplyGlobalStyling(sheet);
        sheet.AutoFilterAll();
        AutoSizeColumns(sheet, orderedColumns);
        sheet.AutoSizeRows();
    }

    void CreateHeadings(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
    {
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            var cell = sheet.Cells[0, i];

            cell.Value = column.Heading;

            ApplyHeadingStyling(cell, column);
        }

        sheet.FreezePanes(1, 0, 1, 0);
    }

    async Task PopulateData(Sheet sheet, List<Column<Style, TModel>> orderedColumns, Cancel cancel)
    {
        //Skip heading
        var startRow = 1;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var index = 0; index < orderedColumns.Count; index++)
            {
                var column = orderedColumns[index];

                var cell = sheet.Cells[xlRow, index];

                var style = cell.GetStyle();
                style.VerticalAlignment = TextAlignmentType.Top;
                style.HorizontalAlignment = TextAlignmentType.Left;
                style.IsTextWrapped = true;
                var value = column.GetValue(item);
                base.SetCellValue(cell, style, value, column, item, trimWhitespace);
                ApplyCellStyle(rowIndex, value, style, column, item);
                cell.SetStyle(style);
            }

            rowIndex++;
        }
    }

    protected override void SetDateFormat(Style style, string format) =>
        style.Custom = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.Custom = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value = value;

    protected override void SetCellHtml(Cell cell, string value) =>
        cell.SafeSetHtml(value);

    protected override void WriteEnumerable(Cell cell, IEnumerable<string> enumerable)
    {
        var value = Html.Enumerable(enumerable, trimWhitespace);
        cell.SafeSetHtml(value);
    }

    void ApplyHeadingStyling(Cell cell, Column<Style, TModel> column)
    {
        var style = cell.GetStyle();
        headingStyle?.Invoke(style);

        column.HeadingStyle?.Invoke(style);

        cell.SetStyle(style);
    }

    void ApplyCellStyle(int index, object? value, Style style, Column<Style, TModel> column, TModel model)
    {
        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.BackgroundColor = alternateRowColor!.Value;
        }

        column.CellStyle?.Invoke(style, model, value);
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

    static void AutoSizeColumns(Sheet sheet, List<Column<Style, TModel>> orderedColumns)
    {
        sheet.AutoSizeColumns();

        for (var index = 0; index < orderedColumns.Count; index++)
        {
            var column = orderedColumns[index];
            if (column.Width != null)
            {
                sheet.Cells.Columns[index].Width = column.Width.Value;
            }
        }
    }
}