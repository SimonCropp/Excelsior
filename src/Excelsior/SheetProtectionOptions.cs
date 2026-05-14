namespace Excelsior;

/// <summary>
/// Defaults match a "data-entry" workbook: the user can edit unlocked data cells,
/// sort, and filter, but cannot change structure or formatting. Booleans that map to
/// SheetProtection attributes use OOXML's "disabled" semantics: true = the action is
/// blocked when the sheet is protected.
/// </summary>
public class SheetProtectionOptions
{
    /// <summary>
    /// If not provided, a fresh GUID is used. That gives a workbook the user can't
    /// accidentally unprotect — fine when the goal is to lock structure rather than
    /// share an unlock code.
    /// </summary>
    public string Password { get; init; } = Guid.NewGuid().ToString();

    /// <summary>
    /// Block editing of embedded objects (shapes, charts, controls).
    /// </summary>
    public bool Objects { get; init; } = true;

    /// <summary>
    /// Block editing of saved scenarios (Data &gt; What-If Analysis &gt; Scenario Manager).
    /// </summary>
    public bool Scenarios { get; init; } = true;

    /// <summary>
    /// Block changing cell formatting (font, fill, number format, etc).
    /// </summary>
    public bool FormatCells { get; init; }

    /// <summary>
    /// Block changing column width / hiding columns.
    /// </summary>
    public bool FormatColumns { get; init; }

    /// <summary>
    /// Block changing row height / hiding rows.
    /// </summary>
    public bool FormatRows { get; init; } = true;

    /// <summary>
    /// Block inserting new columns.
    /// </summary>
    public bool InsertColumns { get; init; } = true;

    /// <summary>
    /// Block inserting new rows.
    /// </summary>
    public bool InsertRows { get; init; } = true;

    /// <summary>
    /// Block inserting hyperlinks.
    /// </summary>
    public bool InsertHyperlinks { get; init; } = true;

    /// <summary>
    /// Block deleting columns.
    /// </summary>
    public bool DeleteColumns { get; init; } = true;

    /// <summary>
    /// Block deleting rows.
    /// </summary>
    public bool DeleteRows { get; init; } = true;

    /// <summary>
    /// Allow selecting locked cells (e.g. headers) so users can copy from them.
    /// </summary>
    public bool SelectLockedCells { get; init; }

    /// <summary>
    /// Allow selecting unlocked (data) cells so they can be edited.
    /// </summary>
    public bool SelectUnlockedCells { get; init; }

    /// <summary>
    /// Allow sorting; useful for reviewing the data.
    /// </summary>
    public bool Sort { get; init; }

    /// <summary>
    /// Allow using the auto-filter dropdowns added to the header row.
    /// </summary>
    public bool AutoFilter { get; init; }

    /// <summary>
    /// Block editing pivot tables and pivot charts.
    /// </summary>
    public bool PivotTables { get; init; } = true;
}
