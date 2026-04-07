namespace Excelsior;

enum CellType { Empty, InlineString, Number, Boolean }

class CellData(CellType type, string? value, uint styleIndex)
{
    public CellType Type => type;
    public string? Value => value;
    public uint StyleIndex { get; private set; } = styleIndex;

    internal void UpdateStyleIndex(uint index) => StyleIndex = index;
}

record HyperlinkInfo(string CellReference, string Url);

class SheetData
{
    public required string Name { get; init; }
    public List<List<CellData>> Rows { get; } = [];
    public List<(int Index, double Width)> ColumnWidths { get; } = [];
    public int FilterFirstColumn { get; set; } = -1;
    public int FilterLastColumn { get; set; } = -1;
    public List<HyperlinkInfo> Hyperlinks { get; } = [];
}
