namespace ExcelsiorOpenXml;

public class OpenXmlColor(string hexValue)
{
    public string HexValue { get; } = hexValue;

    public static OpenXmlColor FromArgb(byte a, byte r, byte g, byte b) =>
        new($"{a:X2}{r:X2}{g:X2}{b:X2}");

    public static OpenXmlColor FromRgb(byte r, byte g, byte b) =>
        FromArgb(255, r, g, b);

    public static OpenXmlColor White => FromRgb(255, 255, 255);
    public static OpenXmlColor Black => FromRgb(0, 0, 0);
    public static OpenXmlColor DarkBlue => FromRgb(0, 0, 139);
    public static OpenXmlColor LightGreen => FromRgb(144, 238, 144);
    public static OpenXmlColor LightPink => FromRgb(255, 182, 193);
    public static OpenXmlColor DarkGreen => FromRgb(0, 100, 0);
}
