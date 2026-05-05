namespace Excelsior;

public class CellAlignment
{
    public HorizontalAlignmentValues Horizontal { get; set; } = HorizontalAlignmentValues.General;
    public VerticalAlignmentValues Vertical { get; set; } = VerticalAlignmentValues.Bottom;
    public bool WrapText { get; set; }
}
