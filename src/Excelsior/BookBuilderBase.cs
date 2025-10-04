namespace Excelsior;

public abstract class BookBuilderBase<TBook, TSheet, TStyle, TCell>
{
    protected abstract TBook BuildBook();

    List<Func<TBook, Cancel, Task>> actions = [];

    public ISheetBuilder<TModel, TStyle> AddSheet<TModel>(
        IEnumerable<TModel> data,
        string? name = null) =>
        AddSheet(data.ToAsyncEnumerable(), name);

    internal abstract SheetRendererBase<TModel, TSheet, TStyle, TCell, TBook> ConstructSheetRenderer<TModel>(
        IAsyncEnumerable<TModel> data,
        string name,
        List<Column<TStyle, TModel>> orderedColumns);

    public ISheetBuilder<TModel, TStyle> AddSheet<TModel>(
        IAsyncEnumerable<TModel> data,
        string? name = null)
    {
        name ??= $"Sheet{actions.Count + 1}";
        var columns = new Columns<TModel, TStyle>();
        var builder = new SheetBuilder<TModel, TStyle>(columns);

        actions.Add((book, cancel) =>
        {
            var renderer = ConstructSheetRenderer(data, name, columns.OrderedColumns());
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