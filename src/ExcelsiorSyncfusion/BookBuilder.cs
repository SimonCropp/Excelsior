namespace ExcelsiorSyncfusion;

public class BookBuilder :
    BookBuilderBase<IDisposableBook, Sheet, Style, Range>
{
    bool useAlternatingRowColors;
    Color? alternateRowColor;
    Action<Style>? headingStyle;
    Action<Style>? globalStyle;
    bool trimWhitespace;
    int defaultMaxCoumnWidth;

    public BookBuilder(
        bool useAlternatingRowColors = false,
        Color? alternateRowColor = null,
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

    internal override RendererBase<TModel, Sheet, Style, Range, IDisposableBook> ConstructSheetRenderer<TModel>(
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
        var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override IDisposableBook BuildBook()
    {
        var engine = new ExcelEngine();
        var book = engine.Excel.Workbooks.Create(0);
        return new DisposableBook(engine, book);
    }
}