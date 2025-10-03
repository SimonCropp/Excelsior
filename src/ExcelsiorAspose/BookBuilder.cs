namespace ExcelsiorAspose;

public class BookBuilder : BookBuilderBase<Book, Sheet,Style>,
IBookBuilder
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

    public override ISheetBuilder<TModel, Style> AddSheet<TModel>(IAsyncEnumerable<TModel> data, string? name = null)
    {
        name ??= $"Sheet{actions.Count + 1}";

        var converter = new SheetBuilder<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace);
        actions.Add((book, cancel) => converter.AddSheet(book, cancel));
        return converter;
    }

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