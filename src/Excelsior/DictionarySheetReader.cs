namespace Excelsior;

class DictionarySheetReader :
    IDictionarySheetReader,
    IReaderSheet
{
    List<ColumnReadInfo> columnInfos = [];
    List<int?> orders = [];
    List<IReadOnlyDictionary<string, object?>> rows = [];
    HashSet<string> names = new(StringComparer.Ordinal);

    public string? Name { get; }
    public IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows => rows;

    public DictionarySheetReader(string? name) =>
        Name = name;

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
        orders.Add(config.Order);

        return this;
    }

    IReadOnlyList<ColumnReadInfo> IReaderSheet.OrderedColumns()
    {
        var indexed = columnInfos
            .Select((info, i) => (info, order: orders[i], declared: i))
            .OrderBy(_ => _.order.HasValue ? 0 : 1)
            .ThenBy(_ => _.order ?? _.declared);

        var result = new List<ColumnReadInfo>(columnInfos.Count);
        foreach (var (info, _, _) in indexed)
        {
            result.Add(info);
        }

        return result;
    }

    void IReaderSheet.Receive(IReadOnlyDictionary<string, object?> rowValues) =>
        rows.Add(rowValues);

    void IReaderSheet.Reset() =>
        rows.Clear();
}
