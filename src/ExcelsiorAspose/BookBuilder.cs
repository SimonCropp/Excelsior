namespace ExcelsiorAspose;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    Color? alternateRowColor = null,
    Action<Style>? headingStyle = null,
    Action<Style>? globalStyle = null,
    bool trimWhitespace = true,
    int defaultMaxColumnWidth = 50) :
    BookBuilderBase<Book, Sheet, Style, Cell, Color?, Column>(
        useAlternatingRowColors,
        alternateRowColor,
        headingStyle,
        globalStyle,
        trimWhitespace,
        defaultMaxColumnWidth)
{
    internal override RendererBase<TModel, Sheet, Style, Cell, Book, Color?, Column> ConstructSheetRenderer<TModel>(
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
        await book.SaveAsync(stream, saveFormat: SaveFormat.Xlsx);
    }

    protected override Book BuildBook()
    {
        var book = new Book();
        book.Worksheets.Clear();
        return book;
    }
}