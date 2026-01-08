using System.Drawing;

namespace ExcelsiorOpenXml;

public class StyleWrapper
{
    public FontProperties Font { get; set; } = new();
    public FillProperties Fill { get; set; } = new();
    public AlignmentProperties Alignment { get; set; } = new();
    public string? NumberFormat { get; set; }
    public uint? NumberFormatId { get; set; }

    public override int GetHashCode() =>
        HashCode.Combine(Font, Fill, Alignment, NumberFormat, NumberFormatId);

    public override bool Equals(object? obj) =>
        obj is StyleWrapper other &&
        Font.Equals(other.Font) &&
        Fill.Equals(other.Fill) &&
        Alignment.Equals(other.Alignment) &&
        NumberFormat == other.NumberFormat &&
        NumberFormatId == other.NumberFormatId;
}

public class FontProperties
{
    public bool Bold { get; set; }
    public System.Drawing.Color? Color { get; set; }
    public string? Name { get; set; }
    public double? Size { get; set; }

    public override int GetHashCode() =>
        HashCode.Combine(Bold, Color, Name, Size);

    public override bool Equals(object? obj) =>
        obj is FontProperties other &&
        Bold == other.Bold &&
        Equals(Color, other.Color) &&
        Name == other.Name &&
        Size == other.Size;
}

public class FillProperties
{
    public System.Drawing.Color? BackgroundColor { get; set; }

    public override int GetHashCode() =>
        BackgroundColor.GetHashCode();

    public override bool Equals(object? obj) =>
        obj is FillProperties other &&
        Equals(BackgroundColor, other.BackgroundColor);
}

public class AlignmentProperties
{
    public HorizontalAlignmentValues? Horizontal { get; set; }
    public VerticalAlignmentValues? Vertical { get; set; }
    public bool WrapText { get; set; }

    public override int GetHashCode() =>
        HashCode.Combine(Horizontal, Vertical, WrapText);

    public override bool Equals(object? obj) =>
        obj is AlignmentProperties other &&
        Horizontal == other.Horizontal &&
        Vertical == other.Vertical &&
        WrapText == other.WrapText;
}
