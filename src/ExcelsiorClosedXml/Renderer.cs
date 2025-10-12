class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    Color? alternateRowColor,
    Action<Style>? headingStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace,
    List<Column<Style, TModel>> columns,
    int maxColumnWidth) :
    RendererBase<TModel, Sheet, Style, Cell, Book>(data, columns, maxColumnWidth)
{
    protected override void ApplyFilter(Sheet sheet) =>
        sheet.RangeUsed()!.SetAutoFilter();

    protected override void FreezeHeader(Sheet sheet) =>
        sheet.SheetView.FreezeRows(1);

    protected override Cell GetCell(Sheet sheet, int row, int column) =>
        sheet.Cell(row + 1, column + 1);

    protected override void ApplyDefaultStyles(Style style)
    {
        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        style.Alignment.WrapText = true;
    }

    protected override Style GetStyle(Cell cell) =>
        cell.Style;

    protected override void ApplyStyle(Cell cell, Style style)
    {
    }

    protected override void RenderCell(object? value, Column<Style, TModel> column, TModel item, int rowIndex, Cell cell)
    {
        var style = cell.Style;
        ApplyDefaultStyles(style);

        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(cell, rowIndex, value, column, item);
    }

    protected override void SetDateFormat(Style style, string format) =>
        style.DateFormat.Format = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat.Format = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value = XLCellValue.FromObject(value);

    protected override void SetCellValue(Cell cell, string value) =>
        cell.SetValue(value);

    protected override void SetCellHtml(Cell cell, string value) =>
        throw new("ClosedXml does not support html");

    protected override Sheet BuildSheet(Book book) =>
        book.Worksheets.Add(name);

    protected override void ApplyHeadingStyling(Cell cell, Column<Style, TModel> column)
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

    protected override void ApplyGlobalStyling(Sheet sheet)
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

    protected override void ResizeColumn(Sheet sheet, int index, Column<Style, TModel> columnConfig, int maxColumnWidth)
    {
        var sheetColumn = sheet.Column(index + 1);
        if (columnConfig.Width == null)
        {
            sheetColumn.AdjustToContents();
            sheetColumn.Width += 2;
            if (sheetColumn.Width > maxColumnWidth)
            {
                sheetColumn.Width = maxColumnWidth;
            }
        }
        else
        {
            sheetColumn.Width = columnConfig.Width.Value;
        }
    }

    protected override void ResizeRows(Sheet sheet) =>
        sheet.Rows().AdjustToContents();
}