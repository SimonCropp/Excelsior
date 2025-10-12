using ExcelsiorAspose;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Cell, Book, Color?, Column>(data, columns, maxColumnWidth, bookBuilder)
{
    protected override void ApplyFilter(Sheet sheet) =>
        sheet.AutoFilterAll();

    protected override void FreezeHeader(Sheet sheet) =>
        sheet.FreezePanes(1, 0, 1, 0);

    protected override Cell GetCell(Sheet sheet, int row, int column) =>
        sheet.Cells[row, column];

    protected override void ApplyDefaultStyles(Style style)
    {
        style.VerticalAlignment = TextAlignmentType.Top;
        style.HorizontalAlignment = TextAlignmentType.Left;
        style.IsTextWrapped = true;
    }

    protected override Style GetStyle(Cell cell) =>
        cell.GetStyle();

    protected override void CommitStyle(Cell cell, Style style) =>
        cell.SetStyle(style);

    protected override void SetStyleColor(Style style, Color? color) =>
        style.BackgroundColor = color!.Value;

    protected override void SetDateFormat(Style style, string format) =>
        style.Custom = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.Custom = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value = value;

    protected override void SetCellValue(Cell cell, string value) =>
        cell.PutValue(value, false);

    protected override void SetCellHtml(Cell cell, string value) =>
        cell.SafeSetHtml(value);

    protected override Sheet BuildSheet(Book book) =>
        book.Worksheets.Add(name);

    protected override void ApplyGlobalStyling(Sheet sheet, Action<Style> globalStyle)
    {
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

    protected override void ResizeColumn(Sheet sheet, int index, ColumnConfig<Style, TModel> columnConfig, int maxColumnWidth)
    {
        var sheetColumn = sheet.Cells.Columns[index];
        if (columnConfig.Width == null)
        {
            sheet.AutoFitColumns(index, index);
            // Round widths since Aspose AutoFitColumns is not deterministic
            sheetColumn.Width = Math.Round(sheetColumn.Width) + 1;

            // Aspose does not seem to respect the dot points
            if (columnConfig.IsEnumerableString)
            {
                sheetColumn.Width += 5;
            }

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
        sheet.AutoSizeRows();
}