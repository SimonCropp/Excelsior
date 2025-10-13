namespace ExcelsiorClosedXml;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    Color? alternateRowColor = null,
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
    internal override RendererBase<TModel, Sheet, Style, Cell, Book, Color, Column> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<ColumnConfig<Style, TModel>> columns,
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

    protected override Book BuildBook() => new XLWorkbook();
}