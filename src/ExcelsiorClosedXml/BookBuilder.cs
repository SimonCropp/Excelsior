namespace ExcelsiorClosedXml;

public class BookBuilder :
    IBookBuilder
{
    List<Func<Book, Cancel, Task>> actions = [];
    bool useAlternatingRowColors;
    XLColor? alternateRowColor;
    Action<IXLStyle>? headerStyle;
    Action<IXLStyle>? globalStyle;
    bool trimWhitespace;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        XLColor? alternateRowColor = null,
        Action<IXLStyle>? headerStyle = null,
        Action<IXLStyle>? globalStyle = null,
        bool trimWhitespace = true)
    {
        ValueRenderer.SetBookBuilderUsed();
        this.useAlternatingRowColors = useAlternatingRowColors;
        this.alternateRowColor = alternateRowColor;
        this.headerStyle = headerStyle;
        this.globalStyle = globalStyle;
        this.trimWhitespace = trimWhitespace;
    }

    public SheetBuilder<TModel> AddSheet<TModel>(IEnumerable<TModel> data, string? name = null)
        where TModel : class =>
        AddSheet(data.ToAsyncEnumerable(), name);

    public SheetBuilder<TModel> AddSheet<TModel>(IAsyncEnumerable<TModel> data, string? name = null)
        where TModel : class
    {
        name ??= $"Sheet{actions.Count + 1}";

        var converter = new SheetBuilder<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headerStyle,
            globalStyle,
            trimWhitespace);
        actions.Add((book, cancel) => converter.AddSheet(book, cancel));
        return converter;
    }

    public async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    public async Task<Book> Build(Cancel cancel = default)
    {
        var book = new XLWorkbook();
        foreach (var action in actions)
        {
            cancel.ThrowIfCancellationRequested();
            await action(book, cancel);
        }

        return book;
    }
}