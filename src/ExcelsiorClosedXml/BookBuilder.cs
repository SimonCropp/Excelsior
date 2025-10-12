namespace ExcelsiorClosedXml;

public class BookBuilder :
    BookBuilderBase<Book, Sheet, Style, Cell, Color>
{
    bool useAlternatingRowColors;
    Color? alternateRowColor;
    Action<Style>? headingStyle;
    Action<Style>? globalStyle;
    bool trimWhitespace;
    int defaultMaxColumnWidth;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        Color? alternateRowColor = null,
        Action<Style>? headingStyle = null,
        Action<Style>? globalStyle = null,
        bool trimWhitespace = true,
        int defaultMaxColumnWidth = 50)
    {
        ValueRenderer.SetBookBuilderUsed();
        this.useAlternatingRowColors = useAlternatingRowColors;
        this.alternateRowColor = alternateRowColor;
        this.headingStyle = headingStyle;
        this.globalStyle = globalStyle;
        this.trimWhitespace = trimWhitespace;
        this.defaultMaxColumnWidth = defaultMaxColumnWidth;
    }

    internal override RendererBase<TModel, Sheet, Style, Cell, Book, Color> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<Column<Style, TModel>> columns,
        int? defaultMaxColumnWidth) =>
        new Renderer<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace,
            columns,
            defaultMaxColumnWidth ?? this.defaultMaxColumnWidth);

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override Book BuildBook() => new XLWorkbook();
}