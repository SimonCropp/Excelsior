namespace Excelsior;

public abstract class BookBuilderBase<TBook, TSheet, TStyle, TCell>
{
    protected List<Func<TBook, Cancel, Task>> actions = [];

    public ISheetBuilder<TModel, TStyle, TCell> AddSheet<TModel>(IEnumerable<TModel> data, string? name = null) =>
        AddSheet(data.ToAsyncEnumerable(), name);

    public abstract ISheetBuilder<TModel, TStyle, TCell> AddSheet<TModel>(IAsyncEnumerable<TModel> data, string? name = null);
}