namespace ExcelsiorOpenXml;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    string? alternateRowColor = null,
    Action<CellStyle>? headingStyle = null,
    Action<CellStyle>? globalStyle = null,
    int defaultMaxColumnWidth = 50) :
    BookBuilderBase<SpreadsheetDocument, SheetContext, CellStyle, Cell, string, ColumnRef>(
        useAlternatingRowColors,
        alternateRowColor,
        headingStyle,
        globalStyle,
        defaultMaxColumnWidth)
{
    internal StyleManager StyleManager { get; } = new();

    internal override RendererBase<TModel, SheetContext, CellStyle, Cell, SpreadsheetDocument, string, ColumnRef>
        ConstructSheetRenderer<TModel>(
            IAsyncEnumerable<TModel> data,
            string name,
            List<ColumnConfig<CellStyle, TModel>> columns,
            int? maxColumnWidth) =>
        new Renderer<TModel>(
            name,
            data,
            columns,
            maxColumnWidth,
            this);

    void ApplyStylesheet(SpreadsheetDocument document)
    {
        var stylesPart = document.WorkbookPart!.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault()
                         ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = StyleManager.BuildStylesheet();
    }

    public override async Task<SpreadsheetDocument> Build(Cancel cancel = default)
    {
        var document = await base.Build(cancel);
        ApplyStylesheet(document);
        return document;
    }

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var document = await Build(cancel);

        if (stream.CanRead)
        {
            document.Clone(stream);
        }
        else
        {
            using var temp = new MemoryStream();
            document.Clone(temp);
            temp.Position = 0;
            temp.CopyTo(stream);
        }
    }

    protected override SpreadsheetDocument BuildBook()
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new(new Sheets());
        return document;
    }
}
