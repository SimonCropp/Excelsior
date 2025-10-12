using ExcelsiorClosedXml;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Cell, Book, Color, Column>(data, columns, maxColumnWidth, bookBuilder)
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

    protected override void CommitStyle(Cell cell, Style style)
    {
    }

    protected override void SetStyleColor(Style style, Color color) =>
        style.Fill.BackgroundColor = color;

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

    protected override void ApplyGlobalStyling(Sheet sheet, Action<Style> globalStyle)
    {
        var lastRow = sheet.LastRowUsed();
        var lastRowNumber = lastRow!.RowNumber();
        var range = sheet.Range(1, 1, lastRowNumber, sheet.LastColumnUsed()!.ColumnNumber());
        globalStyle(range.Style);
    }

    protected override Column GetColumn(Sheet sheet, int index) =>
        sheet.Column(index + 1);

    protected override void SetColumnWidth(Column column, int width) =>
        column.Width = width;

    protected override double GetColumnWidth(Column column) =>
        column.Width;

    protected override void ResizeColumn(Sheet sheet, int index, ColumnConfig<Style, TModel> columnConfig, int maxColumnWidth)
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