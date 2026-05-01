namespace Excelsior;

public sealed record ReadError(
    string SheetName,
    int RowIndex,
    string ColumnName,
    string CellReference,
    string Message,
    Exception? Exception = null)
{
    public override string ToString() =>
        $"{SheetName}!{CellReference} ({ColumnName}, row {RowIndex}): {Message}";
}
