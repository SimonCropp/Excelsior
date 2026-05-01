namespace Excelsior;

public sealed record ReadError(
    string SheetName,
    int RowIndex,
    string ColumnName,
    string CellReference,
    string Message,
    Exception? Inner = null)
{
    public override string ToString() =>
        $"{SheetName}!{CellReference} ({ColumnName}, row {RowIndex}): {Message}";
}
