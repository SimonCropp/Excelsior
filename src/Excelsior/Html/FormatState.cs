namespace Excelsior;

struct FormatState
{
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Underline { get; set; }
    internal bool Strikethrough { get; set; }
    internal string? Color { get; set; }
    internal double? FontSizePt { get; set; }
    internal string? FontFamily { get; set; }
    internal bool Superscript { get; set; }
    internal bool Subscript { get; set; }
    internal string? LinkUrl { get; set; }

    internal readonly bool HasFormatting =>
        Bold ||
        Italic ||
        Underline ||
        Strikethrough ||
        Superscript ||
        Subscript ||
        Color != null ||
        FontSizePt != null ||
        FontFamily != null;
}
