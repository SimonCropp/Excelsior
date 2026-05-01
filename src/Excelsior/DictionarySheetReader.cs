namespace Excelsior;

class DictionarySheetReader(string? name) :
    IDictionarySheetReader,
    IReaderSheet
{
    List<ColumnReadInfo> columnInfos = [];
    List<IReadOnlyDictionary<string, object?>> rows = [];
    HashSet<string> names = new(StringComparer.Ordinal);

    public string? Name { get; } = name;
    public IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows => rows;

    public IDictionarySheetReader Column<TProperty>(
        string name,
        Action<DictionaryColumnReadConfig>? configuration = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Column name must be supplied.", nameof(name));
        }

        if (!names.Add(name))
        {
            throw new($"Sheet already contains a column named '{name}'.");
        }

        var config = new DictionaryColumnReadConfig();
        configuration?.Invoke(config);

        columnInfos.Add(new(
            name,
            config.Heading ?? name,
            typeof(TProperty),
            config.Convert));

        return this;
    }

    IReadOnlyList<ColumnReadInfo> IReaderSheet.Columns() =>
        columnInfos;

    void IReaderSheet.Receive(IReadOnlyDictionary<string, object?> rowValues) =>
        rows.Add(rowValues);

    void IReaderSheet.Reset() =>
        rows.Clear();
}
