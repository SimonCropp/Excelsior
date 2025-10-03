namespace Excelsior;

public abstract class BookBuilderBase<TBook, TSheet, TStyle, TCell>
{
    protected abstract TBook BuildBook();
    protected List<Func<TBook, Cancel, Task>> actions = [];

    public ISheetBuilder<TModel, TStyle, TCell> AddSheet<TModel>(IEnumerable<TModel> data, string? name = null) =>
        AddSheet(data.ToAsyncEnumerable(), name);

    public abstract ISheetBuilder<TModel, TStyle, TCell> AddSheet<TModel>(IAsyncEnumerable<TModel> data, string? name = null);

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
}