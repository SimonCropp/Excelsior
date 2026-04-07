namespace Excelsior;

public class CellStyle
{
    public CellFont Font { get; } = new();
    public string? BackgroundColor { get; set; }
    public CellAlignment Alignment { get; } = new();
    internal string? NumberFormat { get; set; }
}

public class CellFont
{
    public bool Bold { get; set; }
    public bool Underline { get; set; }
    public string? Color { get; set; }
    public double? Size { get; set; }
    public string? Name { get; set; }
}

public class CellAlignment
{
    public HorizontalAlignment Horizontal { get; set; } = HorizontalAlignment.General;
    public VerticalAlignment Vertical { get; set; } = VerticalAlignment.Bottom;
    public bool WrapText { get; set; }
}
