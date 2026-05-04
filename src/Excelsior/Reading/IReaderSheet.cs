interface IReaderSheet
{
    string? Name { get; }
    IReadOnlyList<ColumnReadInfo> Columns();

    /// <summary>
    /// Hand the sheet one row's worth of raw OpenXml cells (in slot order
    /// matching <see cref="Columns"/>). The sheet performs its own conversion
    /// and reports per-cell parse failures via <paramref name="onError"/>
    /// (slot index + message).
    /// </summary>
    void ReceiveRow(Cell?[] cellsBySlot, string?[]? sharedStrings, Action<int, string> onError);

    void Reset();
}
