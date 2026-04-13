namespace Excelsior;

/// <summary>
/// Context passed to a column's formula callback. Exposes the current row's
/// Excel row number and helpers to resolve cell references for other columns.
/// </summary>
public class FormulaContext<TModel>
{
    readonly Dictionary<string, int> columnIndexesByName;

    internal FormulaContext(Dictionary<string, int> columnIndexesByName, int row)
    {
        this.columnIndexesByName = columnIndexesByName;
        Row = row;
    }

    /// <summary>
    /// The 1-based Excel row number of the current data row. The header is
    /// row 1, so the first data row is 2.
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Returns the cell reference (e.g. "B5") for the specified property on
    /// the current row.
    /// </summary>
    public string Ref<TProperty>(Expression<Func<TModel, TProperty>> property) =>
        GetColumnLetter(property) + Row;

    /// <summary>
    /// Returns the column letter (e.g. "B") for the specified property.
    /// </summary>
    public string Column<TProperty>(Expression<Func<TModel, TProperty>> property) =>
        GetColumnLetter(property);

    string GetColumnLetter<TProperty>(Expression<Func<TModel, TProperty>> property)
    {
        var name = property.PropertyName();
        if (!columnIndexesByName.TryGetValue(name, out var columnIndex))
        {
            throw new($"Could not find property in output columns: {name}. Ensure the column is included in the sheet.");
        }

        return SheetContext.GetColumnLetter(columnIndex);
    }
}
