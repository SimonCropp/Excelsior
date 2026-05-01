interface IReaderSheet
{
    string? Name { get; }
    List<ColumnReadInfo> Columns();
    void Receive(IReadOnlyDictionary<string, object?> rowValues);
    void Reset();
}