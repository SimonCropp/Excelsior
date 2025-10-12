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
    RendererBase<TModel, Sheet, Style, Range, IDisposableBook>(data, columns, maxColumnWidth)
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

    protected override void RenderCell(object? value, Column<Style, TModel> column, TModel item, int rowIndex, Range cell, Style style)
    {
        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(rowIndex, value, column, item, style);
    }

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

    protected override void ApplyHeadingStyling(Range cell, Column<Style, TModel> column)
    {
        headingStyle?.Invoke(cell.CellStyle);

        column.HeadingStyle?.Invoke(cell.CellStyle);
    }

    void ApplyCellStyle(int index, object? value, Column<Style, TModel> column, TModel item, Style style)
    {
        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.Color = alternateRowColor!.Value;
        }

        column.CellStyle?.Invoke(style, item, value);
    }

    protected override void ApplyGlobalStyling(Sheet sheet) =>
        globalStyle?.Invoke(sheet.UsedRange.CellStyle);

    protected override void ResizeColumn(Sheet sheet, int index, Column<Style, TModel> columnConfig, int maxColumnWidth)
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