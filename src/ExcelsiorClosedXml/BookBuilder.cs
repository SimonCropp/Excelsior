namespace ExcelsiorClosedXml;

public class BookBuilder :
    BookBuilderBase<Book, Sheet, Style, Cell>
{
    bool useAlternatingRowColors;
    XLColor? alternateRowColor;
    Action<Style>? headingStyle;
    Action<Style>? globalStyle;
    bool trimWhitespace;
    int defaultMaxCoumnWidth;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        XLColor? alternateRowColor = null,
        Action<Style>? headingStyle = null,
        Action<Style>? globalStyle = null,
        bool trimWhitespace = true,
        int defaultMaxCoumnWidth = 50)
    {
        ValueRenderer.SetBookBuilderUsed();
        this.useAlternatingRowColors = useAlternatingRowColors;
        this.alternateRowColor = alternateRowColor;
        this.headingStyle = headingStyle;
        this.globalStyle = globalStyle;
        this.trimWhitespace = trimWhitespace;
        this.defaultMaxCoumnWidth = defaultMaxCoumnWidth;
    }

    internal override RendererBase<TModel, Sheet, Style, Cell, Book> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<Column<Style, TModel>> columns,
        int? defaultMaxCoumnWidth) =>
        new Renderer<TModel>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace,
            columns,
            defaultMaxCoumnWidth ?? this.defaultMaxCoumnWidth);

    public override async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override Book BuildBook() => new XLWorkbook();
}