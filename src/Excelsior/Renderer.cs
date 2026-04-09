class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<TModel>> columns,
    int? maxColumnWidth,
    int? maxRowHeight,
    BookBuilder bookBuilder)
{
    const int MaxExcelRowHeight = 409;
    const int DefaultExcelRowHeight = 15;


    internal bool AutoFilter { get; set; } = true;

    StyleManager? styleManager;
    Dictionary<Cell, CellStyle> cellStyles = [];
    Dictionary<int, double> finalColumnWidths = [];

    internal async Task AddSheet(SpreadsheetDocument book, Cancel cancel)
    {
        var sheet = BuildSheet(book);
        CreateHeadings(sheet);
        FreezeHeader(sheet);
        await PopulateData(sheet, cancel);
        if (bookBuilder.GlobalStyle != null)
        {
            ApplyGlobalStyling(bookBuilder.GlobalStyle);
        }

        var first = -1;
        var last = -1;
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            if (column.IsEnumerable || !(column.Filter ?? AutoFilter))
            {
                continue;
            }

            if (first == -1)
            {
                first = i;
            }

            last = i;
        }

        if (first != -1)
        {
            ApplyFilter(sheet, first, last);
        }

        AutoSizeColumns(sheet);
        ResizeRows(sheet);
        ApplyMaxRowHeight(sheet);
    }

    void CreateHeadings(SheetContext sheet)
    {
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];

            var cell = sheet.GetCell(0, i);

            SetCellValue(cell, column.Heading);
            var style = GetStyle(cell);
            ApplyHeadingStyling(column, style);
            CommitStyle(cell, style);
        }
    }

    void ApplyHeadingStyling(ColumnConfig<TModel> column, CellStyle style)
    {
        style.Font.Bold = true;
        bookBuilder.HeadingStyle?.Invoke(style);
        column.HeadingStyle?.Invoke(style);
    }

    void ResizeColumn(SheetContext sheet, int index, ColumnConfig<TModel> columnConfig)
    {
        var resultMaxColumnWidth = maxColumnWidth ?? bookBuilder.DefaultMaxColumnWidth;
        var column = new ColumnRef(index);
        int width;
        if (columnConfig.Width == null)
        {
            var doubleWidth = AdjustColumnWidth(sheet, column);
            width = (int) Math.Round(doubleWidth);
            width += 1;

            if (columnConfig.IsEnumerable)
            {
                width += 5;
            }

            if (columnConfig.MinWidth is { } min && width < min)
            {
                width = min;
            }

            if (columnConfig.MaxWidth is { } max && width > max)
            {
                width = max;
            }

            if (width > resultMaxColumnWidth)
            {
                width = resultMaxColumnWidth;
            }
        }
        else
        {
            width = columnConfig.Width.Value;
        }

        finalColumnWidths[column.Index] = width;
    }

    void AutoSizeColumns(SheetContext sheet)
    {
        for (var index = 0; index < columns.Count; index++)
        {
            var column = columns[index];
            ResizeColumn(sheet, index, column);
        }
    }

    async Task PopulateData(SheetContext sheet, Cancel cancel)
    {
        var itemIndex = 0;
        await foreach (var item in data.WithCancellation(cancel))
        {
            var rowIndex = itemIndex + 1; // +1 to skip heading;

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var value = column.GetValue(item);
                var cell = sheet.GetCell(rowIndex, columnIndex);
                var style = GetStyle(cell);
                style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
                style.Alignment.Vertical = VerticalAlignmentValues.Top;
                style.Alignment.WrapText = true;
                SetCellValue(cell, sheet, style, value, column, item);

                if (bookBuilder.UseAlternatingRowColors &&
                    rowIndex % 2 == 1)
                {
                    style.BackgroundColor = bookBuilder.AlternateRowColor!;
                }

                column.CellStyle?.Invoke(style, item, value);
                CommitStyle(cell, style);
            }

            itemIndex++;
        }
    }

    CellStyle GetStyle(Cell cell)
    {
        var style = new CellStyle();
        cellStyles[cell] = style;
        return style;
    }

    void CommitStyle(Cell cell, CellStyle style) =>
        cell.StyleIndex = styleManager!.GetOrCreateStyleIndex(style);

    SheetContext BuildSheet(SpreadsheetDocument book)
    {
        styleManager = bookBuilder.StyleManager;

        var workbookPart = book.WorkbookPart!;
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new(new SheetData());

        var sheets = workbookPart.Workbook!.GetFirstChild<Sheets>()!;
        var sheetId = (uint)(sheets.Count() + 1);
        sheets.Append(
            new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            });

        return new(worksheetPart);
    }

    static void FreezeHeader(SheetContext sheet)
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

    static void SetCellValue(Cell cell, object value)
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

    static void SetCellValue(Cell cell, string value)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new(
            new Text(value)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
    }

    static void SetCellHtml(Cell cell, string value) =>
        SpreadsheetHtmlConverter.SetCellHtml(cell, value);

    static void SetCellList(Cell cell, IReadOnlyList<string> items)
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

    static void SetCellLink(Cell cell, SheetContext sheet, CellStyle style, Link link)
    {
        var display = link.Text ?? link.Url;
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new(
            new Text(display)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

        var rel = sheet.WorksheetPart.AddHyperlinkRelationship(new(link.Url), true);
        var hyperlinks = sheet.Worksheet.GetFirstChild<Hyperlinks>();
        if (hyperlinks == null)
        {
            hyperlinks = new();
            sheet.Worksheet.InsertAfter(hyperlinks, sheet.SheetData);
        }

        hyperlinks.Append(
            new Hyperlink
            {
                Reference = cell.CellReference,
                Id = rel.Id
            });

        style.Font.Color = "0563C1";
        style.Font.Underline = true;
    }

    static void SetCellLinkList(Cell cell, SheetContext sheet, List<string> items, string? hyperlinkUrl)
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
                    new RunProperties(
                        new Underline(),
                        new Color
                        {
                            Rgb = "0563C1"
                        }),
                    new Text(items[i])
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
        }

        cell.InlineString = inlineString;

        if (hyperlinkUrl == null)
        {
            return;
        }

        var rel = sheet.WorksheetPart.AddHyperlinkRelationship(new(hyperlinkUrl), true);
        var hyperlinks = sheet.Worksheet.GetFirstChild<Hyperlinks>();
        if (hyperlinks == null)
        {
            hyperlinks = new();
            sheet.Worksheet.InsertAfter(hyperlinks, sheet.SheetData);
        }

        hyperlinks.Append(
            new Hyperlink
            {
                Reference = cell.CellReference,
                Id = rel.Id
            });
    }

    void ApplyGlobalStyling(Action<CellStyle> globalStyle)
    {
        foreach (var (cell, style) in cellStyles)
        {
            globalStyle(style);
            cell.StyleIndex = styleManager!.GetOrCreateStyleIndex(style);
        }
    }

    static void ApplyFilter(SheetContext sheet, int firstColumn, int lastColumn)
    {
        if (sheet.RowCount == 0)
        {
            return;
        }

        var firstCol = SheetContext.GetColumnLetter(firstColumn);
        var lastCol = SheetContext.GetColumnLetter(lastColumn);
        var reference = $"{firstCol}1:{lastCol}{sheet.RowCount}";
        sheet.Worksheet
            .InsertAfter(
                new AutoFilter
                {
                    Reference = reference
                },
                sheet.SheetData);
    }

    static double AdjustColumnWidth(SheetContext sheet, ColumnRef column)
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

    void ResizeRows(SheetContext sheet)
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

    void ApplyMaxRowHeight(SheetContext sheet)
    {
        var max = maxRowHeight ?? bookBuilder.MaxRowHeight;
        if (max == null)
        {
            return;
        }

        if (max < DefaultExcelRowHeight || max > MaxExcelRowHeight)
        {
            throw new($"MaxRowHeight ({max}) must be between {DefaultExcelRowHeight} (the Excel default row height) and {MaxExcelRowHeight}.");
        }

        var maxLinesAllowed = max.Value / (double)DefaultExcelRowHeight;

        foreach (var row in sheet.SheetData.Elements<Row>())
        {
            double maxLines = 1;
            for (var i = 0; i < columns.Count; i++)
            {
                var colLetter = SheetContext.GetColumnLetter(i);
                var cellRef = colLetter + row.RowIndex;
                var cell = row.Elements<Cell>()
                    .FirstOrDefault(_ => _.CellReference?.Value == cellRef);
                if (cell == null)
                {
                    continue;
                }

                var width = finalColumnWidths.GetValueOrDefault(i, 8d);
                var lines = EstimateVisualLines(cell, width);
                if (lines > maxLines)
                {
                    maxLines = lines;
                }
            }

            if (maxLines > maxLinesAllowed)
            {
                row.Height = (double)max;
                row.CustomHeight = true;
            }
        }
    }

    static double EstimateVisualLines(Cell cell, double columnWidth)
    {
        var charsPerLine = Math.Max(1d, (columnWidth - 2) / 1.1);
        double lines = 0;
        var hasContent = false;
        foreach (var text in EnumerateCellTexts(cell))
        {
            hasContent = true;
            foreach (var line in text.Split('\n'))
            {
                lines += Math.Max(1d, Math.Ceiling(line.Length / charsPerLine));
            }
        }

        return hasContent ? lines : 1;
    }

    static IEnumerable<string> EnumerateCellTexts(Cell cell)
    {
        if (cell.InlineString != null)
        {
            var hasRuns = false;
            foreach (var run in cell.InlineString.Elements<Run>())
            {
                hasRuns = true;
                if (run.Text?.Text is { } runText)
                {
                    yield return runText;
                }
            }

            if (!hasRuns && cell.InlineString.Text?.Text is { } inlineText)
            {
                yield return inlineText;
            }

            yield break;
        }

        if (cell.CellValue?.Text is { } cellValueText)
        {
            yield return cellValueText;
        }
    }

    static void SetCellValue(
        Cell cell,
        SheetContext sheet,
        CellStyle style,
        object? value,
        ColumnConfig<TModel> column,
        TModel item)
    {
        void SetStringOrHtml(string content)
        {
            if (column.IsHtml)
            {
                SetCellHtml(cell, content);
            }
            else
            {
                SetCellValue(cell, content);
            }
        }

        void ThrowIfHtml()
        {
            if (column.IsHtml)
            {
                throw new("TreatAsHtml is not compatible with this type");
            }
        }

        if (value == null)
        {
            if (column.NullDisplay != null)
            {
                SetCellValue(cell, column.NullDisplay);
            }

            return;
        }

        if (column.TryRender(item, value, out var render))
        {
            SetStringOrHtml(render);

            return;
        }

        if (value is Link link)
        {
            SetCellLink(cell, sheet, style, link);
            return;
        }

        if (column.IsEnumerable &&
            value is IEnumerable<Link?> linkEnumerable)
        {
            var links = new List<Link>();
            foreach (var l in linkEnumerable)
            {
                if (l == null)
                {
                    continue;
                }

                links.Add(l);
            }

            if (links.Count > 0)
            {
                var linkItems = new List<string>(links.Count);
                foreach (var l in links)
                {
                    linkItems.Add(l.Text == null ? l.Url : $"{l.Text} ({l.Url})");
                }

                var hyperlinkUrl = links.Count == 1 ? links[0].Url : null;
                SetCellLinkList(cell, sheet, linkItems, hyperlinkUrl);
            }

            return;
        }

        if (column.IsEnumerable &&
            value is IEnumerable enumerable)
        {
            var items = new List<string>();
            foreach (var obj in enumerable)
            {
                if (obj == null)
                {
                    continue;
                }

                var str = column.ItemRender == null ? obj.ToString() : column.ItemRender(obj);
                if (str != null &&
                    ValueRenderer.TrimWhitespace)
                {
                    str = str.Trim();
                }

                if (str != null)
                {
                    items.Add(str);
                }
            }

            if (items.Count > 0)
            {
                SetCellList(cell, items);
            }

            return;
        }

        if (value is DateTime dateTime)
        {
            ThrowIfHtml();
            style.NumberFormat = column.Format ?? ValueRenderer.DefaultDateTimeFormat;
            SetCellValue(cell, dateTime);

            return;
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            ThrowIfHtml();

            var format = column.Format ?? ValueRenderer.DefaultDateTimeOffsetFormat;
            style.NumberFormat = format;
            SetCellValue(cell, dateTimeOffset.ToString(format, CultureInfo.InvariantCulture));

            return;
        }

        if (value is Date date)
        {
            ThrowIfHtml();
            style.NumberFormat = column.Format ?? ValueRenderer.DefaultDateFormat;
            SetCellValue(cell, date.ToDateTime(new(0, 0)));

            return;
        }

        if (value is bool boolean)
        {
            ThrowIfHtml();
            SetCellValue(cell, boolean);
            return;
        }

        if (column.IsNumber)
        {
            ThrowIfHtml();
            if (column.Format != null)
            {
                style.NumberFormat = column.Format;
            }

            SetCellValue(cell, Convert.ToDouble(value));
            return;
        }

        var valueAsString = value.ToString();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
        if (valueAsString != null &&
            ValueRenderer.TrimWhitespace)
        {
            valueAsString = valueAsString.Trim();
        }

        if (valueAsString != null)
        {
            SetStringOrHtml(valueAsString);
        }
    }
}
