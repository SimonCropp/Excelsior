class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    XLColor? alternateRowColor,
    Action<Style>? headingStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace,
    List<Column<Style, TModel>> columns) :
    RendererBase<TModel, Sheet, Style, Cell, Book>(data, columns)
{
    internal override async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        CreateHeadings(sheet);

        await PopulateData(sheet, cancel);

        ApplyGlobalStyling(sheet);
        sheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(sheet);
    }

    void CreateHeadings(Sheet sheet)
    {
        for (var i = 0; i < Columns.Count; i++)
        {
            var column = Columns[i];
            var cell = sheet.Cell(1, i + 1);

            cell.Value = column.Heading;

            ApplyHeadingStyling(cell, column);
        }

        sheet.SheetView.FreezeRows(1);
    }


    protected override Cell GetCell(Sheet sheet, int row, int column) =>
        sheet.Cell(row + 1, column + 1);

    protected override void RenderCell(object? value,
        Column<Style, TModel> column,
        TModel item,
        int rowIndex,
        Cell cell)
    {
        var style = cell.Style;
        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        style.Alignment.WrapText = true;

        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(cell, rowIndex, value, column, item);
    }

    protected override void SetDateFormat(Style style, string format) =>
        style.DateFormat.Format = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat.Format = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value = XLCellValue.FromObject(value);

    protected override void SetCellHtml(Cell cell, string value) =>
        throw new("ClosedXml does not support html");

    protected override void WriteEnumerable(Cell cell, IEnumerable<string> enumerable) =>
        RichText.Enumerable(cell, enumerable, trimWhitespace);

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

    void ApplyGlobalStyling(Sheet sheet)
    {
        if (globalStyle == null)
        {
            return;
        }

        var lastRow = sheet.LastRowUsed();
        var lastRowNumber = lastRow!.RowNumber();
        var range = sheet.Range(1, 1, lastRowNumber, Columns.Count);
        globalStyle(range.Style);
    }

    void AutoSizeColumns(Sheet sheet)
    {
        var endRow = sheet.RowsUsed().Count();
        for (var index = 0; index < Columns.Count; index++)
        {
            var columnConfig = Columns[index];
            var column = sheet.Column(index + 1);
            if (columnConfig.Width == null)
            {
                column.AdjustToContents(1, endRow);
                column.Width += 2;
            }
            else
            {
                column.Width = columnConfig.Width.Value;
            }
        }
    }
}