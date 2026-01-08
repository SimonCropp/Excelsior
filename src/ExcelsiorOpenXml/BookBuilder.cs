using System.Drawing;

namespace ExcelsiorOpenXml;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    Color alternateRowColor = null,
    Action<Style>? headingStyle = null,
    Action<Style>? globalStyle = null,
    int defaultMaxColumnWidth = 50) :
    BookBuilderBase<Book, Sheet, Style, Cell, Color, Column>(
        useAlternatingRowColors,
        alternateRowColor,
        headingStyle,
        globalStyle,
        defaultMaxColumnWidth)
{
    MemoryStream? documentStream;

    protected override Book BuildBook()
    {
        documentStream = new MemoryStream();
        var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook, autoSave: true);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook
        {
            Sheets = new Sheets()
        };
        return document;
    }

    internal override RendererBase<TModel, Sheet, Style, Cell, Book, Color, Column> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<ColumnConfig<Style, TModel>> columns,
        int? maxColumnWidth) =>
        new Renderer<TModel>(name, data, columns, maxColumnWidth, this);

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        var document = await Build(cancel);

        // Dispose the document to flush everything to the underlying stream
        document.Dispose();

        // Copy the memory stream to output
        if (documentStream != null)
        {
            documentStream.Position = 0;
            await documentStream.CopyToAsync(stream, cancel);
        }
    }
}
