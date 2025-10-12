namespace Excelsior;

public abstract class BookBuilderBase<TBook, TSheet, TStyle, TCell, TColor>
{
    protected BookBuilderBase(
        bool useAlternatingRowColors,
        TColor? alternateRowColor,
        Action<TStyle>? headingStyle,
        Action<TStyle>? globalStyle,
        bool trimWhitespace,
        int defaultMaxColumnWidth)
    {
        ValueRenderer.SetBookBuilderUsed();
        UseAlternatingRowColors = useAlternatingRowColors;
        AlternateRowColor = alternateRowColor;
        HeadingStyle = headingStyle;
        GlobalStyle = globalStyle;
        TrimWhitespace = trimWhitespace;
        DefaultMaxColumnWidth = defaultMaxColumnWidth;
    }

    public bool UseAlternatingRowColors { get; }

    protected abstract TBook BuildBook();

    List<Func<TBook, Cancel, Task>> actions = [];
    public int DefaultMaxColumnWidth{ get; }
    public TColor? AlternateRowColor{ get; }
    public Action<TStyle>? HeadingStyle{ get; }
    public Action<TStyle>? GlobalStyle{ get; }
    public bool TrimWhitespace{ get; }

    public ISheetBuilder<TModel, TStyle> AddSheet<TModel>(
        IEnumerable<TModel> data,
        string? name = null,
        int? defaultMaxColumnWidth = null) =>
        AddSheet(data.ToAsyncEnumerable(), name, defaultMaxColumnWidth);

    internal abstract RendererBase<TModel, TSheet, TStyle, TCell, TBook, TColor> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<Column<TStyle, TModel>> columns,
        int? defaultMaxColumnWidth);

    public ISheetBuilder<TModel, TStyle> AddSheet<TModel>(
        IAsyncEnumerable<TModel> data,
        string? name = null,
        int? defaultMaxColumnWidth = null)
    {
        name ??= $"Sheet{actions.Count + 1}";
        var columns = new Columns<TModel, TStyle>();
        var builder = new SheetBuilder<TModel, TStyle>(columns);

        actions.Add((book, cancel) =>
        {
            var renderer = ConstructSheetRenderer(data, name, columns.OrderedColumns(), defaultMaxColumnWidth);
            return renderer.AddSheet(book, cancel);
        });

        return builder;
    }

    public async Task<TBook> Build(Cancel cancel = default)
    {
        var book = BuildBook();
        foreach (var action in actions)
        {
            cancel.ThrowIfCancellationRequested();
            await action(book, cancel);
        }

        return book;
    }

    public abstract Task ToStream(Stream stream, Cancel cancel = default);

    public async Task ToFile(string path, Cancel cancel = default)
    {
        await using var stream = File.Create(path);
        await ToStream(stream, cancel);
    }
}