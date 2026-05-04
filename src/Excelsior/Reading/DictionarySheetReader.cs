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

    public IReadOnlyList<ColumnReadInfo> Columns() =>
        columnInfos;

    public void ReceiveRow(Cell?[] cellsBySlot, string?[]? sharedStrings, Action<int, string> onError)
    {
        var dict = new Dictionary<string, object?>(cellsBySlot.Length, StringComparer.Ordinal);
        for (var slot = 0; slot < cellsBySlot.Length; slot++)
        {
            var cell = cellsBySlot[slot];
            var column = columnInfos[slot];

            if (column.Convert != null && cell != null)
            {
                try
                {
                    dict[column.Name] = column.Convert(cell);
                }
                catch (Exception exception)
                {
                    onError(slot, $"Converter delegate threw: {exception.Message}");
                }

                continue;
            }

            if (CellConverter.TryConvert(cell, column.Type, sharedStrings, out var value, out var error))
            {
                dict[column.Name] = value;
            }
            else
            {
                onError(slot, error);
            }
        }

        rows.Add(dict);
    }

    public void Reset() =>
        rows.Clear();
}
