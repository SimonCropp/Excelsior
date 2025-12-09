namespace ExcelsiorOpenXml;

public class CellWrapper(Cell cell)
{
    public Cell Cell { get; } = cell;
    public OpenXmlStyle Style { get; } = new();
}
