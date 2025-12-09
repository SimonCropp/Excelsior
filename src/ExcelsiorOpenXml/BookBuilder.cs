using DocumentFormat.OpenXml;

namespace ExcelsiorOpenXml;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    OpenXmlColor? alternateRowColor = null,
    Action<OpenXmlStyle>? headingStyle = null,
    Action<OpenXmlStyle>? globalStyle = null,
    int defaultMaxColumnWidth = 50) :
        BookBuilderBase<Book, Sheet, OpenXmlStyle, CellWrapper, OpenXmlColor, OpenXmlColumn>(
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            defaultMaxColumnWidth)
{
    MemoryStream? stream;

    internal override RendererBase<TModel, Sheet, OpenXmlStyle, CellWrapper, Book, OpenXmlColor, OpenXmlColumn> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<ColumnConfig<OpenXmlStyle, TModel>> columns,
        int? maxColumnWidth) =>
        new Renderer<TModel>(
            name,
            data,
            columns,
            maxColumnWidth,
            this);

    public override async Task ToStream(Stream targetStream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.Clone(targetStream);
    }

    protected override Book BuildBook()
    {
        stream = new();
        var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new SharedStringTable();

        return document;
    }

    static Stylesheet CreateStylesheet()
    {
        var stylesheet = new Stylesheet
        {
            Fonts = new Fonts(new Font()),
            Fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            Borders = new Borders(new Border()),
            CellFormats = new CellFormats(new CellFormat())
        };

        stylesheet.Fonts.Count = 1;
        stylesheet.Fills.Count = 2;
        stylesheet.Borders.Count = 1;
        stylesheet.CellFormats.Count = 1;

        return stylesheet;
    }
}
