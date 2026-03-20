using System.Text;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Cell, IDisposableBook, Color?, Column>(data, columns, maxColumnWidth, bookBuilder)
{
    protected override void ApplyFilter(Sheet sheet, int firstColumn, int lastColumn)
    {
        var lastRow = sheet.UsedRange.LastRow;
        sheet.AutoFilters.FilterRange = sheet.Range[1, firstColumn + 1, lastRow, lastColumn + 1];
    }

    protected override void FreezeHeader(Sheet sheet) =>
        sheet.Rows[0].FreezePanes();

    protected override Cell GetCell(Sheet sheet, int row, int column) =>
        sheet.Range[row + 1, column + 1];

    protected override void ApplyDefaultStyles(Style style)
    {
        style.HorizontalAlignment = ExcelHAlign.HAlignLeft;
        style.VerticalAlignment = ExcelVAlign.VAlignTop;
        style.WrapText = true;
    }

    protected override Style GetStyle(Cell cell) =>
        cell.CellStyle;

    protected override void CommitStyle(Cell cell, Style style)
    {
    }

    protected override void SetStyleColor(Style style, Color? color) =>
        style.Color = color!.Value;

    protected override void SetDateFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetCellValue(Cell cell, object value) =>
        cell.Value2 = value;

    protected override void SetCellValue(Cell cell, string value) =>
        cell.Text = value;

    protected override void SetCellHtml(Cell cell, string value) =>
        cell.HtmlString = value;

    protected override void SetCellList(Cell cell, IReadOnlyList<string> items)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                builder.Append('\n');
            }

            builder.Append("● ");
            builder.Append(items[i]);
        }

        cell.Text = builder.ToString();

        var pos = 0;
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                pos++; // newline
            }

            var font = cell.Worksheet.Workbook.CreateFont();
            font.Bold = true;
            cell.RichText.SetFont(pos, pos + 1, font);
            pos += 2 + items[i].Length;
        }
    }

    protected override void SetCellLink(Cell cell, Sheet sheet, Style style, Link link)
    {
        var display = link.Text ?? link.Url;
        cell.Worksheet.HyperLinks.Add(cell, ExcelHyperLinkType.Url, link.Url, display);
        style.Font.Color = ExcelKnownColors.Blue;
        style.Font.Underline = ExcelUnderline.Single;
    }

    protected override void SetCellLinkList(Cell cell, Sheet sheet, Style style, IReadOnlyList<string> items, string? hyperlinkUrl)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                builder.Append('\n');
            }

            builder.Append("● ");
            builder.Append(items[i]);
        }

        cell.Text = builder.ToString();

        var pos = 0;
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                pos++; // newline
            }

            var bulletFont = cell.Worksheet.Workbook.CreateFont();
            bulletFont.Bold = true;
            bulletFont.Color = ExcelKnownColors.Blue;
            bulletFont.Underline = ExcelUnderline.Single;
            cell.RichText.SetFont(pos, pos + 1, bulletFont);
            var textFont = cell.Worksheet.Workbook.CreateFont();
            textFont.Color = ExcelKnownColors.Blue;
            textFont.Underline = ExcelUnderline.Single;
            cell.RichText.SetFont(pos + 2, pos + 1 + items[i].Length, textFont);
            pos += 2 + items[i].Length;
        }

        if (hyperlinkUrl != null)
        {
            cell.Worksheet.HyperLinks.Add(cell, ExcelHyperLinkType.Url, hyperlinkUrl, cell.Text);
        }
    }

    protected override void SetBold(Style style) =>
        style.Font.Bold = true;

    protected override Sheet BuildSheet(IDisposableBook book) =>
        book.Worksheets.Create(name);

    protected override void ApplyGlobalStyling(Sheet sheet, Action<Style> globalStyle) =>
        globalStyle.Invoke(sheet.UsedRange.CellStyle);

    protected override Column GetColumn(Sheet sheet, int index) =>
        sheet.Columns[index];

    protected override void SetColumnWidth(Column column, int width) =>
        column.ColumnWidth = width;

    protected override double AdjustColumnWidth(Sheet sheet, Column column)
    {
        column.AutofitColumns();
        return column.ColumnWidth;
    }

    protected override void ResizeRows(Sheet sheet)
    {
        sheet.UsedRange.AutofitRows();

        // Clamp row heights to Excel's maximum of 409 points
        var lastRow = sheet.UsedRange.LastRow;
        for (var i = 1; i <= lastRow; i++)
        {
            var height = sheet.Range[i, 1].RowHeight;
            if (height > 409)
            {
                sheet.Range[i, 1].RowHeight = 409;
            }
        }
    }
}