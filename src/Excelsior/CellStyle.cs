namespace Excelsior;

public class CellStyle
{
    public CellFont Font { get; } = new();
    public CellFill Fill { get; } = new();
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

public class CellFill
{
    public string? BackgroundColor { get; set; }
}

public class CellAlignment
{
    public HorizontalAlignmentValues Horizontal { get; set; } = HorizontalAlignmentValues.General;
    public VerticalAlignmentValues Vertical { get; set; } = VerticalAlignmentValues.Bottom;
    public bool WrapText { get; set; }
}
