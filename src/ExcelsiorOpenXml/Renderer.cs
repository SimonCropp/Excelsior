using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<OpenXmlStyle, TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder) :
    RendererBase<TModel, Sheet, OpenXmlStyle, CellWrapper, Book, OpenXmlColor, OpenXmlColumn>(data, columns, maxColumnWidth, bookBuilder)
{
    WorksheetPart? worksheetPart;
    SheetData? sheetData;
    readonly Dictionary<string, CellWrapper> cellCache = [];
    readonly Dictionary<uint, OpenXmlColumn> columnCache = [];
    readonly int columnCount = columns.Count;

    protected override void ApplyFilter(Sheet sheet)
    {
        if (sheetData == null || worksheetPart == null)
        {
            return;
        }

        var lastRow = sheetData.Elements<Row>().LastOrDefault();
        if (lastRow?.RowIndex?.Value == null)
        {
            return;
        }

        var lastColumnIndex = columnCount;
        var range = $"A1:{GetColumnLetter(lastColumnIndex - 1)}{lastRow.RowIndex.Value}";

        var autoFilter = new AutoFilter { Reference = range };
        sheet.InsertAfter(autoFilter, sheetData);
    }

    protected override void FreezeHeader(Sheet sheet)
    {
        var sheetViews = sheet.GetFirstChild<SheetViews>();
        if (sheetViews == null)
        {
            sheetViews = new SheetViews();
            sheet.InsertAt(sheetViews, 0);
        }

        var sheetView = sheetViews.GetFirstChild<SheetView>();
        if (sheetView == null)
        {
            sheetView = new SheetView { WorkbookViewId = 0 };
            sheetViews.AppendChild(sheetView);
        }

        var pane = new Pane
        {
            VerticalSplit = 1,
            TopLeftCell = "A2",
            ActivePane = PaneValues.BottomLeft,
            State = PaneStateValues.Frozen
        };

        sheetView.AppendChild(pane);
    }

    protected override CellWrapper GetCell(Sheet sheet, int row, int column)
    {
        var cellReference = $"{GetColumnLetter(column)}{row + 1}";

        if (cellCache.TryGetValue(cellReference, out var cached))
        {
            return cached;
        }

        if (sheetData == null)
        {
            throw new InvalidOperationException("SheetData not initialized");
        }

        var rowIndex = (uint)(row + 1);
        var rowElement = sheetData.Elements<Row>().FirstOrDefault(_ => _.RowIndex?.Value == rowIndex);
        if (rowElement == null)
        {
            rowElement = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(rowElement);
        }

        var cell = new Cell { CellReference = cellReference };
        rowElement.AppendChild(cell);

        var wrapper = new CellWrapper(cell);
        cellCache[cellReference] = wrapper;

        return wrapper;
    }

    protected override void ApplyDefaultStyles(OpenXmlStyle style)
    {
        style.Alignment.Horizontal = OpenXmlStyle.HorizontalAlignment.Left;
        style.Alignment.Vertical = OpenXmlStyle.VerticalAlignment.Top;
        style.Alignment.WrapText = true;
    }

    protected override OpenXmlStyle GetStyle(CellWrapper cell) =>
        cell.Style;

    protected override void CommitStyle(CellWrapper cellWrapper, OpenXmlStyle style)
    {
        // Styles in OpenXML are applied via StyleIndex, which requires building a stylesheet
        // For simplicity, we'll apply basic formatting directly to cells where possible
        // More complex styling would require extending the stylesheet
    }

    protected override void SetStyleColor(OpenXmlStyle style, OpenXmlColor color) =>
        style.Fill.BackgroundColor = color;

    protected override void SetDateFormat(OpenXmlStyle style, string format) =>
        style.DateFormat = format;

    protected override void SetNumberFormat(OpenXmlStyle style, string format) =>
        style.NumberFormat = format;

    protected override void SetCellValue(CellWrapper cellWrapper, object value)
    {
        var cell = cellWrapper.Cell;

        switch (value)
        {
            case bool boolValue:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolValue);
                break;
            case double or float or decimal or int or long or short:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDouble(value));
                break;
            case DateTime dateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dateTime.ToOADate());
                cell.StyleIndex = 0; // Would need proper date style
                break;
            default:
                SetCellValue(cellWrapper, value.ToString() ?? string.Empty);
                break;
        }
    }

    protected override void SetCellValue(CellWrapper cellWrapper, string value)
    {
        var cell = cellWrapper.Cell;

        if (worksheetPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is { } sharedStringPart)
        {
            var index = InsertSharedStringItem(value, sharedStringPart);
            cell.DataType = CellValues.SharedString;
            cell.CellValue = new CellValue(index.ToString());
        }
        else
        {
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(value);
        }
    }

    protected override void SetCellHtml(CellWrapper cell, string value) =>
        throw new NotSupportedException("OpenXML does not support HTML in cells");

    protected override Sheet BuildSheet(Book book)
    {
        var workbookPart = book.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");

        worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var worksheet = new Worksheet();
        sheetData = new SheetData();
        worksheet.AppendChild(sheetData);
        worksheetPart.Worksheet = worksheet;

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? throw new InvalidOperationException("Sheets is null");
        var sheetId = (uint) (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() + 1);

        var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name
        };

        sheets.AppendChild(sheet);

        return worksheet;
    }

    protected override void ApplyGlobalStyling(Sheet sheet, Action<OpenXmlStyle> globalStyle)
    {
        // Global styling in OpenXML requires updating the stylesheet
        // This is a simplified implementation
    }

    protected override OpenXmlColumn GetColumn(Sheet sheet, int index)
    {
        var columnIndex = (uint) (index + 1);

        if (columnCache.TryGetValue(columnIndex, out var cached))
        {
            return cached;
        }

        var columns = sheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            sheet.InsertBefore(columns, sheetData);
        }

        var column = new DocumentFormat.OpenXml.Spreadsheet.Column
        {
            Min = columnIndex,
            Max = columnIndex,
            CustomWidth = true
        };

        columns.AppendChild(column);

        var wrapper = new OpenXmlColumn(column);
        columnCache[columnIndex] = wrapper;

        return wrapper;
    }

    protected override void SetColumnWidth(OpenXmlColumn column, int width)
    {
        column.Width = width;
        column.Column.Width = width;
    }

    protected override double AdjustColumnWidth(Sheet sheet, OpenXmlColumn column) =>
        // OpenXML doesn't have auto-sizing built-in
        // Return a default width that will be adjusted by the base class
        column.Width ?? 10;

    protected override void ResizeRows(Sheet sheet)
    {
        // Row auto-sizing in OpenXML would require measuring text
        // This is typically handled by Excel when opening the file
    }

    static string GetColumnLetter(int columnIndex)
    {
        var columnLetter = string.Empty;
        var modulo = columnIndex;

        while (modulo >= 0)
        {
            columnLetter = (char) ('A' + (modulo % 26)) + columnLetter;
            modulo = modulo / 26 - 1;
        }

        return columnLetter;
    }

    static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        var sharedStringTable = shareStringPart.SharedStringTable;

        var index = 0;
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return index;
            }

            index++;
        }

        sharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        sharedStringTable.Save();

        return index;
    }
}
