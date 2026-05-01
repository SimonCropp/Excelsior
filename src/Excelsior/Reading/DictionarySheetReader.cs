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
        Func<Cell, TProperty>? convert = null)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Column name must be supplied.", nameof(name));
        }

        if (!names.Add(name))
        {
            throw new($"Sheet already contains a column named '{name}'.");
        }

        Func<Cell, object?>? boxed = convert == null ? null : cell => convert(cell);
        columnInfos.Add(new(
            name,
            name,
            typeof(TProperty),
            boxed));

        return this;
    }

    public List<ColumnReadInfo> Columns() =>
        columnInfos;

    public void Receive(IReadOnlyDictionary<string, object?> rowValues) =>
        rows.Add(rowValues);

    public void Reset() =>
        rows.Clear();
}
