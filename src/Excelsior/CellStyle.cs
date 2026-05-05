namespace Excelsior;

public class CellStyle
{
    public CellFont Font { get; } = new();
    public string? BackgroundColor { get; set; }
    public CellAlignment Alignment { get; } = new();
    public bool Locked { get; set; } = true;
    internal string? NumberFormat { get; set; }
}
