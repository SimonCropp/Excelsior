namespace ExcelsiorSyncfusion;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    Color? alternateRowColor = null,
    Action<Style>? headingStyle = null,
    Action<Style>? globalStyle = null,
    bool trimWhitespace = true,
    int defaultMaxColumnWidth = 50) :
        BookBuilderBase<IDisposableBook, Sheet, Style, Range, Color?>(
            useAlternatingRowColors,
            alternateRowColor,
            headingStyle,
            globalStyle,
            trimWhitespace,
            defaultMaxColumnWidth)
{
    internal override RendererBase<TModel, Sheet, Style, Range, IDisposableBook, Color?> ConstructSheetRenderer<TModel>(
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
        var book = await Build(cancel);
        book.SaveAs(stream);
    }

    protected override IDisposableBook BuildBook()
    {
        var engine = new ExcelEngine
        {
            Excel =
            {
                DefaultVersion = ExcelVersion.Excel2016
            }
        };
        var book = engine.Excel.Workbooks.Create(0);
        book.DetectDateTimeInValue = false;
        return new DisposableBook(engine, book);
    }
}