namespace ExcelsiorOpenXml;

public class OpenXmlColor(string hexValue)
{
    public string HexValue { get; } = hexValue;

    public static Color FromArgb(byte a, byte r, byte g, byte b) =>
        new($"{a:X2}{r:X2}{g:X2}{b:X2}");

    public static Color FromRgb(byte r, byte g, byte b) =>
        FromArgb(255, r, g, b);

    public static Color White => FromRgb(255, 255, 255);
    public static Color Black => FromRgb(0, 0, 0);
    public static Color DarkBlue => FromRgb(0, 0, 139);
    public static Color LightGreen => FromRgb(144, 238, 144);
    public static Color LightPink => FromRgb(255, 182, 193);
    public static Color DarkGreen => FromRgb(0, 100, 0);
}
