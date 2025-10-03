namespace ExcelsiorClosedXml;

public class BookBuilder :
    BookBuilderBase<Book, Sheet,Style, Cell>,
    IBookBuilder
{
    bool useAlternatingRowColors;
    XLColor? alternateRowColor;
    Action<Style>? headingStyle;
    Action<Style>? globalStyle;
    bool trimWhitespace;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        XLColor? alternateRowColor = null,
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

        var converter = ConstructSheetBuilder(data, name);
        actions.Add((book, cancel) => converter.AddSheet(book, cancel));
        return converter;
    }

    protected override SheetBuilderBase<TModel, Style, Cell, Book> ConstructSheetBuilder<TModel>(IAsyncEnumerable<TModel> data, string name) =>
        new SheetBuilder<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace);

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override Book BuildBook() => new XLWorkbook();
}