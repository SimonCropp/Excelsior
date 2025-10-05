namespace ExcelsiorAspose;

public class BookBuilder : BookBuilderBase<Book, Sheet, Style, Cell>
{
    bool useAlternatingRowColors;
    Color? alternateRowColor;
    Action<Style>? headingStyle;
    Action<Style>? globalStyle;
    bool trimWhitespace;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        Color? alternateRowColor = null,
        Action<Style>? headingStyle = null,
        Action<Style>? globalStyle = null,
        bool trimWhitespace = true)
    {
        ValueRenderer.SetBookBuilderUsed();
        this.useAlternatingRowColors = useAlternatingRowColors;
        this.alternateRowColor = alternateRowColor;
        this.headingStyle = headingStyle;
        this.globalStyle = globalStyle;
        this.trimWhitespace = trimWhitespace;
    }

    internal override RendererBase<TModel, Sheet, Style, Cell, Book> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<Column<Style, TModel>> columns) =>
        new Renderer<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace,
            columns);

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