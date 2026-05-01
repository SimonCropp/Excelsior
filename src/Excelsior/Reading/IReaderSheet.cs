interface IReaderSheet
{
    string? Name { get; }
    IReadOnlyList<ColumnReadInfo> Columns();
    void Receive(IReadOnlyDictionary<string, object?> rowValues);
    void Reset();
}