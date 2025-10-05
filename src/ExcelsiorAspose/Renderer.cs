class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    bool useAlternatingRowColors,
    Color? alternateRowColor,
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
        sheet.AutoFilterAll();
        AutoSizeColumns(sheet);
        sheet.AutoSizeRows();
    }

    void CreateHeadings(Sheet sheet)
    {
        for (var i = 0; i < Columns.Count; i++)
        {
            var column = Columns[i];
            var cell = sheet.Cells[0, i];

            cell.Value = column.Heading;

            ApplyHeadingStyling(cell, column);
        }

        sheet.FreezePanes(1, 0, 1, 0);
    }

    protected override Cell GetCell(Sheet sheet, int row, int column) =>
        sheet.Cells[row, column];

    protected override void RenderCell(object? value, Column<Style, TModel> column, TModel item, int rowIndex, Cell cell)
    {
        var style = cell.GetStyle();
        style.VerticalAlignment = TextAlignmentType.Top;
        style.HorizontalAlignment = TextAlignmentType.Left;
        style.IsTextWrapped = true;
        base.SetCellValue(cell, style, value, column, item, trimWhitespace);
        ApplyCellStyle(rowIndex, value, style, column, item);
        cell.SetStyle(style);
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

    protected override void ResizeColumn(Sheet sheet, int index, int? width)
    {
        var column = sheet.Cells.Columns[index];
        if (width == null)
        {
            sheet.AutoFitColumns(index, index);
            //Round widths since aspose AutoFitColumns is not deterministic
            column.Width = Math.Round(column.Width) + 1;
        }
        else
        {
            column.Width = width.Value;
        }
    }
}