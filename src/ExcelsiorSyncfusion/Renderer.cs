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
    internal override async Task AddSheet(IDisposableBook book, Cancel cancel)
    {
        var sheet = book.Worksheets.Create(name);

        CreateHeadings(sheet);

        await PopulateData(sheet, cancel);

        ApplyGlobalStyling(sheet);
        sheet.AutoFilters.FilterRange = sheet.UsedRange;
        AutoSizeColumns(sheet);
    }

    void CreateHeadings(Sheet sheet)
    {
        for (var i = 0; i < Columns.Count; i++)
        {
            var column = Columns[i];
            var cell = sheet.Range[1, i + 1];

            cell.Value = column.Heading;

            ApplyHeadingStyling(cell, column);
        }

        sheet.Rows[0].FreezePanes();
    }

    protected override Range GetCell(Sheet sheet, int row, int column) =>
        sheet.Range[row + 1, column + 1];

    protected override void RenderCell(object? value,
        Column<Style, TModel> column,
        TModel item,
        int rowIndex,
        Range cell)
    {
        var style = cell.CellStyle;
        style.HorizontalAlignment = ExcelHAlign.HAlignLeft;
        style.VerticalAlignment = ExcelVAlign.VAlignTop;
        style.WrapText = true;

        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(cell, rowIndex, value, column, item);
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

    void ApplyHeadingStyling(Range cell, Column<Style, TModel> column)
    {
        headingStyle?.Invoke(cell.CellStyle);

        column.HeadingStyle?.Invoke(cell.CellStyle);
    }

    void ApplyCellStyle(Range cell, int index, object? value, Column<Style, TModel> column, TModel item)
    {
        var style = cell.CellStyle;

        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.Color = alternateRowColor!.Value;
        }

        column.CellStyle?.Invoke(style, item, value);
    }

    void ApplyGlobalStyling(Sheet sheet) =>
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
}