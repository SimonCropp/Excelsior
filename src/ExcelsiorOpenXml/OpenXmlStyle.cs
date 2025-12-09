namespace ExcelsiorOpenXml;

public class OpenXmlStyle
{
    public FontStyle Font { get; } = new();
    public FillStyle Fill { get; } = new();
    public AlignmentStyle Alignment { get; } = new();
    public string? DateFormat { get; set; }
    public string? NumberFormat { get; set; }

    public class FontStyle
    {
        public bool Bold { get; set; }
        public OpenXmlColor? FontColor { get; set; }
        public double? FontSize { get; set; }
        public string? FontName { get; set; }
    }

    public class FillStyle
    {
        public OpenXmlColor? BackgroundColor { get; set; }
    }

    public class AlignmentStyle
    {
        public HorizontalAlignment? Horizontal { get; set; }
        public VerticalAlignment? Vertical { get; set; }
        public bool WrapText { get; set; }
    }

    public enum HorizontalAlignment
    {
        Left,
        Center,
        Right
    }

    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    }
}
