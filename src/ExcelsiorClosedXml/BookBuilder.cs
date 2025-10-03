namespace ExcelsiorClosedXml;

public class BookBuilder : BookBuilderBase<Book, Sheet,Style,Cell>,
    IBookBuilder
{
    List<Func<Book, Cancel, Task>> actions = [];
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

    public override ISheetBuilder<TModel, Style, Cell> AddSheet<TModel>(IAsyncEnumerable<TModel> data, string? name = null)
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