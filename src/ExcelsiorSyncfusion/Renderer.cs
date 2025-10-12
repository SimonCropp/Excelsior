class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Range, IDisposableBook, Color?, Range>(data, columns, maxColumnWidth, bookBuilder)
{
    protected override void ApplyFilter(Sheet sheet) =>
        sheet.AutoFilters.FilterRange = sheet.UsedRange;

    protected override void FreezeHeader(Sheet sheet) =>
        sheet.Rows[0].FreezePanes();

    protected override Range GetCell(Sheet sheet, int row, int column) =>
        sheet.Range[row + 1, column + 1];

    protected override void ApplyDefaultStyles(Style style)
    {
        style.HorizontalAlignment = ExcelHAlign.HAlignLeft;
        style.VerticalAlignment = ExcelVAlign.VAlignTop;
        style.WrapText = true;
    }

    protected override Style GetStyle(Range cell) =>
        cell.CellStyle;

    protected override void CommitStyle(Range cell, Style style)
    {
    }

    protected override void SetStyleColor(Style style, Color? color) =>
        style.Color = color!.Value;

    protected override void SetDateFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetCellValue(Range cell, object value) =>
        cell.Value2 = value;

    protected override void SetCellValue(Range cell, string value) =>
        cell.Text = value;

    protected override void SetCellHtml(Range cell, string value) =>
        cell.HtmlString = value;

    protected override Sheet BuildSheet(IDisposableBook book) =>
        book.Worksheets.Create(name);

    protected override void ApplyGlobalStyling(Sheet sheet, Action<Style> globalStyle) =>
        globalStyle.Invoke(sheet.UsedRange.CellStyle);

    protected override Range GetColumn(Sheet sheet, int index) =>
        sheet.Columns[index];

    protected override void SetColumnWidth(Range column, int width) =>
        column.ColumnWidth = width;

    protected override double GetColumnWidth(Range column) =>
         column.ColumnWidth;

    protected override void ResizeColumn(Sheet sheet, int index, ColumnConfig<Style, TModel> columnConfig, int maxColumnWidth)
    {
        var sheetColumn = sheet.Columns[index];
        if (columnConfig.Width == null)
        {
            sheet.AutofitColumn(index + 1);
            sheetColumn.ColumnWidth += 4;

            // does not seem to respect the dot points
            if (columnConfig.IsEnumerableString)
            {
                sheetColumn.ColumnWidth += 5;
            }

            if (sheetColumn.ColumnWidth > maxColumnWidth)
            {
                sheetColumn.ColumnWidth = maxColumnWidth;
            }
        }
        else
        {
            sheetColumn.ColumnWidth = columnConfig.Width.Value;
        }
    }

    protected override void ResizeRows(Sheet sheet) =>
        sheet.UsedRange.AutofitRows();
}