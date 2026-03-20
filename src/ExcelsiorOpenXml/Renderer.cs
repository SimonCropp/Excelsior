class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<CellStyle, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, SheetContext, CellStyle, Cell, OpenXmlBook, string, ColumnRef>(
        data, columns, maxColumnWidth, bookBuilder)
{
    StyleManager? styleManager;
    Dictionary<Cell, CellStyle> cellStyles = [];
    Dictionary<int, double> finalColumnWidths = [];

    protected override void SetBold(CellStyle style) =>
        style.Font.Bold = true;

    protected override SheetContext BuildSheet(OpenXmlBook book)
    {
        styleManager = book.StyleManager;

        var workbookPart = book.Document.WorkbookPart!;
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new(new SheetData());

        var sheets = workbookPart.Workbook!.GetFirstChild<Sheets>()!;
        var sheetId = (uint)(sheets.Count() + 1);
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name
        });

        return new(worksheetPart);
    }

    protected override void FreezeHeader(SheetContext sheet)
    {
        var sheetViews = new SheetViews(
            new SheetView(
                new Pane
                {
                    VerticalSplit = 1,
                    TopLeftCell = "A2",
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen
                })
            {
                TabSelected = true,
                WorkbookViewId = 0
            });
        sheet.Worksheet.InsertBefore(sheetViews, sheet.SheetData);
    }

    protected override Cell GetCell(SheetContext sheet, int row, int column) =>
        sheet.GetCell(row, column);

    protected override void ApplyDefaultStyles(CellStyle style)
    {
        style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
        style.Alignment.Vertical = VerticalAlignmentValues.Top;
        style.Alignment.WrapText = true;
    }

    protected override CellStyle GetStyle(Cell cell)
    {
        var style = new CellStyle();
        cellStyles[cell] = style;
        return style;
    }

    protected override void CommitStyle(Cell cell, CellStyle style) =>
        cell.StyleIndex = styleManager!.GetOrCreateStyleIndex(style);

    protected override void SetStyleColor(CellStyle style, string color) =>
        style.Fill.BackgroundColor = color;

    protected override void SetDateFormat(CellStyle style, string format) =>
        style.NumberFormat = format;

    protected override void SetNumberFormat(CellStyle style, string format) =>
        style.NumberFormat = format;

    protected override void SetCellValue(Cell cell, object value)
    {
        switch (value)
        {
            case bool b:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new(b);
                break;
            case DateTime dt:
                cell.CellValue = new(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                break;
            case double d:
                cell.CellValue = new(d.ToString(CultureInfo.InvariantCulture));
                break;
            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new(
                    new Text(value.ToString() ?? "")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                break;
        }
    }

    protected override void SetCellValue(Cell cell, string value)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new(
            new Text(value)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
    }

    protected override void SetCellHtml(Cell cell, string value) =>
        SpreadsheetHtmlConverter.SetCellHtml(cell, value);

    protected override void SetCellList(Cell cell, IReadOnlyList<string> items)
    {
        cell.DataType = CellValues.InlineString;
        var inlineString = new InlineString();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                inlineString.Append(
                    new Run(
                        new Text("\n")
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        }));
            }

            inlineString.Append(
                new Run(
                    new RunProperties(new Bold()),
                    new Text("● ")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
            inlineString.Append(
                new Run(
                    new Text(items[i])
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
        }

        cell.InlineString = inlineString;
    }

    protected override void SetCellLink(Cell cell, SheetContext sheet, CellStyle style, Link link)
    {
        var display = link.Text ?? link.Url;
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new(new Text(display) { Space = SpaceProcessingModeValues.Preserve });

        var rel = sheet.WorksheetPart.AddHyperlinkRelationship(new Uri(link.Url), true);
        var hyperlinks = sheet.Worksheet.GetFirstChild<Hyperlinks>();
        if (hyperlinks == null)
        {
            hyperlinks = new Hyperlinks();
            sheet.Worksheet.InsertAfter(hyperlinks, sheet.SheetData);
        }

        hyperlinks.Append(new Hyperlink { Reference = cell.CellReference, Id = rel.Id });

        style.Font.Color = "0563C1";
        style.Font.Underline = true;
    }

    protected override void SetCellLinkList(Cell cell, SheetContext sheet, CellStyle style, IReadOnlyList<string> items, string? hyperlinkUrl)
    {
        cell.DataType = CellValues.InlineString;
        var blueColor = new Color { Rgb = "0563C1" };
        var inlineString = new InlineString();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                inlineString.Append(
                    new Run(
                        new Text("\n")
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        }));
            }

            inlineString.Append(
                new Run(
                    new RunProperties(new Bold(), new Underline(), blueColor.CloneNode(true)),
                    new Text("● ")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
            inlineString.Append(
                new Run(
                    new RunProperties(new Underline(), blueColor.CloneNode(true)),
                    new Text(items[i])
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
        }

        cell.InlineString = inlineString;

        if (hyperlinkUrl != null)
        {
            var rel = sheet.WorksheetPart.AddHyperlinkRelationship(new Uri(hyperlinkUrl), true);
            var hyperlinks = sheet.Worksheet.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null)
            {
                hyperlinks = new Hyperlinks();
                sheet.Worksheet.InsertAfter(hyperlinks, sheet.SheetData);
            }

            hyperlinks.Append(new Hyperlink { Reference = cell.CellReference, Id = rel.Id });
        }
    }

    protected override void ApplyGlobalStyling(SheetContext sheet, Action<CellStyle> globalStyle)
    {
        foreach (var (cell, style) in cellStyles)
        {
            globalStyle(style);
            cell.StyleIndex = styleManager!.GetOrCreateStyleIndex(style);
        }
    }

    protected override void ApplyFilter(SheetContext sheet, int firstColumn, int lastColumn)
    {
        if (sheet.RowCount == 0)
        {
            return;
        }

        var firstCol = SheetContext.GetColumnLetter(firstColumn);
        var lastCol = SheetContext.GetColumnLetter(lastColumn);
        var reference = $"{firstCol}1:{lastCol}{sheet.RowCount}";
        sheet.Worksheet.InsertAfter(new AutoFilter
        {
            Reference = reference
        }, sheet.SheetData);
    }

    protected override ColumnRef GetColumn(SheetContext sheet, int index) =>
        new(index);

    protected override void SetColumnWidth(ColumnRef column, int width) =>
        finalColumnWidths[column.Index] = width;

    protected override double AdjustColumnWidth(SheetContext sheet, ColumnRef column)
    {
        double maxWidth = 8;
        var colLetter = SheetContext.GetColumnLetter(column.Index);

        foreach (var row in sheet.SheetData.Elements<Row>())
        {
            var cellRef = colLetter + row.RowIndex;
            var cell = row.Elements<Cell>()
                .FirstOrDefault(_ => _.CellReference?.Value == cellRef);
            if (cell == null)
            {
                continue;
            }

            var length = GetCellContentLength(cell);
            var estimated = length * 1.1 + 2;
            if (estimated > maxWidth)
            {
                maxWidth = estimated;
            }
        }

        return maxWidth;
    }

    static int GetCellContentLength(Cell cell)
    {
        if (cell.InlineString != null)
        {
            var length = 0;
            var hasRuns = false;
            foreach (var run in cell.InlineString.Elements<Run>())
            {
                hasRuns = true;
                length += run.Text?.Text.Length ?? 0;
            }

            if (hasRuns)
            {
                return length;
            }

            if (cell.InlineString.Text != null)
            {
                return cell.InlineString.Text.Text.Length;
            }
        }

        return cell.CellValue?.Text.Length ?? 0;
    }

    protected override void ResizeRows(SheetContext sheet)
    {
        if (finalColumnWidths.Count <= 0)
        {
            return;
        }

        var cols = new Columns();
        foreach (var (index, width) in finalColumnWidths.OrderBy(_ => _.Key))
        {
            cols.Append(
                new Column
                {
                    Min = (uint)(index + 1),
                    Max = (uint)(index + 1),
                    Width = width,
                    CustomWidth = true
                });
        }

        sheet.Worksheet.InsertBefore(cols, sheet.SheetData);
    }
}
