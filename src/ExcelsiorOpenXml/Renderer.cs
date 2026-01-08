using System.Drawing;

namespace ExcelsiorOpenXml;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<Style, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, Style, Cell, Book, Color, Column>(data, columns, maxColumnWidth, bookBuilder)
{
    readonly Dictionary<int, ColumnMetadata> columnMetadata = new();
    readonly Dictionary<Style, uint> styleCache = new();
    readonly Dictionary<int, double> columnMaxWidths = new(); // Track max width per column
    SheetData? sheetData;
    WorkbookStylesPart? stylesPart;
    uint nextStyleIndex = 1; // Start at 1, index 0 is reserved for default format

    protected override Cell GetCell(Sheet sheet, int row, int column)
    {
        // Ensure SheetData exists
        if (sheetData == null)
        {
            sheetData = sheet.Worksheet.GetFirstChild<SheetData>() ?? sheet.Worksheet.AppendChild(new SheetData());
        }

        // Get or create row
        var rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)(row + 1));
        if (rowElement == null)
        {
            rowElement = new Row { RowIndex = (uint)(row + 1) };
            sheetData.Append(rowElement);
        }

        // Create cell
        var cellReference = GetCellReference(row, column);
        var cell = new Cell
        {
            CellReference = cellReference,
            DataType = CellValues.String
        };
        rowElement.Append(cell);

        return cell;
    }

    static string GetCellReference(int row, int column)
    {
        var columnName = "";
        var index = column;
        while (index >= 0)
        {
            columnName = (char)('A' + index % 26) + columnName;
            index = index / 26 - 1;
        }
        return $"{columnName}{row + 1}";
    }

    protected override Style GetStyle(Cell cell) => new();

    protected override void CommitStyle(Cell cell, Style style)
    {
        // Get or create style index
        if (!styleCache.TryGetValue(style, out var styleIndex))
        {
            styleIndex = nextStyleIndex++;
            styleCache[style] = styleIndex;

            // Add style to stylesheet (will be built at the end)
        }

        cell.StyleIndex = styleIndex;
    }

    protected override void ApplyDefaultStyles(Style style)
    {
        style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
        style.Alignment.Vertical = VerticalAlignmentValues.Top;
        style.Alignment.WrapText = true;
    }

    protected override void SetCellValue(Cell cell, object value)
    {
        cell.DataType = value switch
        {
            bool => CellValues.Boolean,
            int or long or short or byte or decimal or double or float => CellValues.Number,
            DateTime => CellValues.Number, // DateTime stored as number with format
            _ => CellValues.String
        };

        // Convert DateTime to Excel's OLE Automation date format
        if (value is DateTime dateTime)
        {
            cell.CellValue = new CellValue(dateTime.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        else
        {
            cell.CellValue = new CellValue(value.ToString() ?? "");
        }

        // Track column width
        TrackColumnWidth(cell, value);
    }

    protected override void SetCellValue(Cell cell, string value)
    {
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue(value);

        // Track column width
        TrackColumnWidth(cell, value);
    }

    void TrackColumnWidth(Cell cell, object? value)
    {
        var columnIndex = GetColumnIndexFromReference(cell.CellReference?.Value);
        if (columnIndex < 0) return;

        var estimatedWidth = WidthEstimator.EstimateWidth(value);

        if (!columnMaxWidths.TryGetValue(columnIndex, out var currentMax) || estimatedWidth > currentMax)
        {
            columnMaxWidths[columnIndex] = estimatedWidth;
        }
    }

    static int GetColumnIndexFromReference(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference)) return -1;

        // Extract column letters (e.g., "A1" -> "A", "AB23" -> "AB")
        var columnLetters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        if (string.IsNullOrEmpty(columnLetters)) return -1;

        // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, etc.)
        var index = 0;
        foreach (var ch in columnLetters)
        {
            index = index * 26 + (ch - 'A' + 1);
        }
        return index - 1;
    }

    protected override void SetCellHtml(Cell cell, string value)
    {
        try
        {
            // Convert HTML to runs with formatting
            var runs = HtmlConverter.ConvertToRuns(value);

            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString();

            foreach (var run in runs)
            {
                cell.InlineString.Append(run);
            }

            // Track column width based on stripped text
            TrackColumnWidth(cell, value);
        }
        catch
        {
            // Fallback to plain text if HTML conversion fails
            SetCellValue(cell, value);
        }
    }

    protected override void SetDateFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetNumberFormat(Style style, string format) =>
        style.NumberFormat = format;

    protected override void SetStyleColor(Style style, Color color)
    {
        if (color.HasValue)
        {
            style.Fill.BackgroundColor = color.Value;
        }
    }

    protected override Sheet BuildSheet(Book book)
    {
        var worksheetPart = book.WorkbookPart!.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet();
        sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

        // Create or get stylesheet
        stylesPart = book.WorkbookPart.WorkbookStylesPart ?? book.WorkbookPart.AddNewPart<WorkbookStylesPart>();

        // Add sheet to workbook
        var sheets = book.WorkbookPart.Workbook.Sheets!;
        var sheetId = (uint)(sheets.Count() + 1);
        sheets.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet
        {
            Id = book.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name
        });

        return worksheetPart;
    }

    protected override Column GetColumn(Sheet sheet, int index)
    {
        if (!columnMetadata.TryGetValue(index, out var column))
        {
            column = new ColumnMetadata
            {
                Index = index,
                Name = GetCellReference(0, index).TrimEnd('1'),
                Width = 10
            };
            columnMetadata[index] = column;
        }
        return column;
    }

    protected override void SetColumnWidth(Column column, int width) =>
        column.Width = width;

    protected override double AdjustColumnWidth(Sheet sheet, Column column)
    {
        // Return the tracked maximum width for this column
        if (columnMaxWidths.TryGetValue(column.Index, out var maxWidth))
        {
            return maxWidth;
        }

        // Default width if no cells in this column
        return column.Width;
    }

    protected override void FreezeHeader(Sheet sheet)
    {
        var sheetViews = sheet.Worksheet.GetFirstChild<SheetViews>() ?? sheet.Worksheet.InsertAt(new SheetViews(), 0);
        var sheetView = sheetViews.GetFirstChild<SheetView>() ?? sheetViews.AppendChild(new SheetView { WorkbookViewId = 0 });

        sheetView.Pane = new Pane
        {
            VerticalSplit = 1,
            TopLeftCell = "A2",
            ActivePane = PaneValues.BottomLeft,
            State = PaneStateValues.Frozen
        };
    }

    protected override void ApplyFilter(Sheet sheet)
    {
        if (sheetData == null || !sheetData.Elements<Row>().Any() || columnMetadata.Count == 0) return;

        var lastRow = sheetData.Elements<Row>().Max(r => r.RowIndex!.Value);
        var lastCol = columnMetadata.Keys.Max();
        var range = $"A1:{GetCellReference((int)lastRow - 1, lastCol)}";

        sheet.Worksheet.AppendChild(new AutoFilter { Reference = range });
    }

    protected override void ApplyGlobalStyling(Sheet sheet, Action<Style> globalStyle)
    {
        if (sheetData == null) return;

        // Iterate through all cells and apply global style
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                // Get current style or create new one
                var style = new Style();

                // Apply default styles first
                ApplyDefaultStyles(style);

                // Apply global style
                globalStyle(style);

                // Commit the updated style
                CommitStyle(cell, style);
            }
        }
    }

    protected override void ResizeRows(Sheet sheet)
    {
        // Build stylesheet with all collected styles
        BuildStylesheet();

        // Apply column widths
        ApplyColumnWidths(sheet);
    }

    void BuildStylesheet()
    {
        if (stylesPart == null) return;

        // Build fonts
        var fonts = new Fonts { Count = 0 };
        var fontCache = new Dictionary<FontProperties, uint>();
        fonts.Append(new Font()); // Default font at index 0
        fonts.Count++;

        // Build fills
        var fills = new Fills { Count = 0 };
        var fillCache = new Dictionary<FillProperties, uint>();
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.None })); // Required default
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.Gray125 })); // Required default
        fills.Count = 2;

        // Build borders
        var borders = new Borders(new Border()) { Count = 1 };

        // Build number formats
        var numberingFormats = new NumberingFormats();
        var formatCache = new Dictionary<string, uint>();
        uint nextFormatId = 164; // Custom formats start at 164

        // Build cell formats
        var cellFormats = new CellFormats { Count = 0 };
        cellFormats.Append(new CellFormat()); // Default format at index 0
        cellFormats.Count++;

        // Process each unique style
        foreach (var (style, styleIndex) in styleCache.OrderBy(x => x.Value))
        {
            // Get or create font
            uint fontId = 0;
            if (style.Font.Bold || style.Font.Color.HasValue)
            {
                if (!fontCache.TryGetValue(style.Font, out fontId))
                {
                    fontId = fonts.Count!.Value;
                    var font = new Font();

                    if (style.Font.Bold)
                        font.Append(new Bold());

                    if (style.Font.Color.HasValue)
                    {
                        font.Append(new DocumentFormat.OpenXml.Spreadsheet.Color
                        {
                            Rgb = new HexBinaryValue(
                                $"{style.Font.Color.Value.A:X2}{style.Font.Color.Value.R:X2}" +
                                $"{style.Font.Color.Value.G:X2}{style.Font.Color.Value.B:X2}")
                        });
                    }

                    fonts.Append(font);
                    fonts.Count++;
                    fontCache[style.Font] = fontId;
                }
            }

            // Get or create fill
            uint fillId = 0;
            if (style.Fill.BackgroundColor.HasValue)
            {
                if (!fillCache.TryGetValue(style.Fill, out fillId))
                {
                    fillId = fills.Count!.Value;
                    var fill = new Fill(
                        new PatternFill(
                            new ForegroundColor
                            {
                                Rgb = new HexBinaryValue(
                                    $"{style.Fill.BackgroundColor.Value.A:X2}{style.Fill.BackgroundColor.Value.R:X2}" +
                                    $"{style.Fill.BackgroundColor.Value.G:X2}{style.Fill.BackgroundColor.Value.B:X2}")
                            })
                        {
                            PatternType = PatternValues.Solid
                        });

                    fills.Append(fill);
                    fills.Count++;
                    fillCache[style.Fill] = fillId;
                }
            }

            // Get or create number format
            uint? numberFormatId = null;
            if (style.NumberFormat != null)
            {
                if (!formatCache.TryGetValue(style.NumberFormat, out var formatId))
                {
                    formatId = nextFormatId++;
                    numberingFormats.Append(new NumberingFormat
                    {
                        NumberFormatId = formatId,
                        FormatCode = style.NumberFormat
                    });
                    formatCache[style.NumberFormat] = formatId;
                }
                numberFormatId = formatId;
            }

            // Create cell format
            var cellFormat = new CellFormat
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = 0,
                ApplyFont = fontId > 0,
                ApplyFill = fillId > 0,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Horizontal = style.Alignment.Horizontal,
                    Vertical = style.Alignment.Vertical,
                    WrapText = style.Alignment.WrapText
                }
            };

            if (numberFormatId.HasValue)
            {
                cellFormat.NumberFormatId = numberFormatId.Value;
                cellFormat.ApplyNumberFormat = true;
            }

            cellFormats.Append(cellFormat);
            cellFormats.Count++;
        }

        // Build stylesheet
        var stylesheet = new Stylesheet();

        if (numberingFormats.Any())
        {
            numberingFormats.Count = (uint)numberingFormats.Count();
            stylesheet.Append(numberingFormats);
        }

        stylesheet.Append(fonts);
        stylesheet.Append(fills);
        stylesheet.Append(borders);
        stylesheet.Append(cellFormats);

        stylesPart.Stylesheet = stylesheet;
    }

    void ApplyColumnWidths(Sheet sheet)
    {
        if (columnMetadata.Count == 0) return;

        var columns = new Columns();
        foreach (var col in columnMetadata.Values.OrderBy(c => c.Index))
        {
            columns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column
            {
                Min = (uint)(col.Index + 1),
                Max = (uint)(col.Index + 1),
                Width = col.Width,
                CustomWidth = true
            });
        }

        sheet.Worksheet.InsertAt(columns, 0);
    }
}
