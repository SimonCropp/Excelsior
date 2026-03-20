using System.Text;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Cell, Book, Color?, Column>(data, columns, maxColumnWidth, bookBuilder)
{
    protected override void ApplyFilter(Sheet sheet, int firstColumn, int lastColumn)
    {
        var cells = sheet.Cells;
        var row = CellsHelper.RowIndexToName(cells.MaxRow);
        var firstCol = CellsHelper.ColumnIndexToName(firstColumn);
        var lastCol = CellsHelper.ColumnIndexToName(lastColumn);
        sheet.AutoFilter.Range = $"{firstCol}1:{lastCol}{row}";
    }

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

        cell.PutValue(builder.ToString());

        var pos = 0;
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                pos++; // newline
            }

            cell.Characters(pos, 2).Font.IsBold = true;
            pos += 2 + items[i].Length;
        }
    }

    protected override void SetCellLink(Cell cell, Sheet sheet, Style style, Link link)
    {
        cell.PutValue(link.Text ?? link.Url, false);
        sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, link.Url);
        style.Font.Color = Color.Blue;
        style.Font.Underline = FontUnderlineType.Single;
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

        cell.PutValue(builder.ToString());

        var pos = 0;
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                pos++; // newline
            }

            var bulletChars = cell.Characters(pos, 2);
            bulletChars.Font.IsBold = true;
            bulletChars.Font.Color = Color.Blue;
            bulletChars.Font.Underline = FontUnderlineType.Single;
            var textChars = cell.Characters(pos + 2, items[i].Length);
            textChars.Font.Color = Color.Blue;
            textChars.Font.Underline = FontUnderlineType.Single;
            pos += 2 + items[i].Length;
        }

        if (hyperlinkUrl != null)
        {
            sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, hyperlinkUrl);
        }
    }

    protected override void SetBold(Style style) =>
        style.Font.IsBold = true;

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

    protected override Column GetColumn(Sheet sheet, int index) =>
        sheet.Cells.Columns[index];

    protected override void SetColumnWidth(Column column, int width) =>
        column.Width = width;

    protected override double AdjustColumnWidth(Sheet sheet,Column column)
    {
        sheet.AutoFitColumns(column.Index, column.Index);
        return column.Width;
    }

    protected override void ResizeRows(Sheet sheet) =>
        sheet.AutoSizeRows();
}