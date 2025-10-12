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

    protected override void ApplyStyle(Cell cell, Style style) =>
        cell.SetStyle(style);

    protected override void RenderCell(object? value, Column<Style, TModel> column, TModel item, int rowIndex, Cell cell, Style style)
    {
        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(rowIndex, value, style, column, item);
    }

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

    protected override void ApplyHeadingStyling(Cell cell, Column<Style, TModel> column)
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

    protected override void ApplyGlobalStyling(Sheet sheet)
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

    protected override void ResizeColumn(Sheet sheet, int index, Column<Style, TModel> columnConfig, int maxColumnWidth)
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