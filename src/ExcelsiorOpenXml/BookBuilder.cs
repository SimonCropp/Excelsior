namespace ExcelsiorOpenXml;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    string? alternateRowColor = null,
    Action<CellStyle>? headingStyle = null,
    Action<CellStyle>? globalStyle = null,
    int defaultMaxColumnWidth = 50) :
    BookBuilderBase<OpenXmlBook, SheetContext, CellStyle, Cell, string, ColumnRef>(
        useAlternatingRowColors,
        alternateRowColor,
        headingStyle,
        globalStyle,
        defaultMaxColumnWidth)
{
    internal override RendererBase<TModel, SheetContext, CellStyle, Cell, OpenXmlBook, string, ColumnRef>
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

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override OpenXmlBook BuildBook() => new();
}
