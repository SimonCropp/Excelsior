using System.Xml;

namespace Excelsior;

class StyleManager
{
    record FontKey(bool Bold, bool Underline, string? Color, double? Size, string? Name);

    record CellFormatKey(
        uint FontId,
        uint FillId,
        uint? NumberFormatId,
        HorizontalAlignment HAlign,
        VerticalAlignment VAlign,
        bool WrapText);

    List<FontKey> fonts = [new(false, false, null, null, null)];

    Dictionary<FontKey, uint> fontIndex = new()
    {
        [new(false, false, null, null, null)] = 0
    };

    List<string?> fills = [null, "gray125"];

    Dictionary<string, uint> fillIndex = [];

    List<string> customNumberFormats = [];
    Dictionary<string, uint> numberFormatIds = [];
    uint nextNumberFormatId = 164;

    List<CellFormatKey> cellFormats =
        [new(0, 0, null, HorizontalAlignment.General, VerticalAlignment.Bottom, false)];

    Dictionary<CellFormatKey, uint> cellFormatIndex = [];

    internal uint GetOrCreateStyleIndex(CellStyle style)
    {
        var fontId = GetOrCreateFontId(style.Font);
        var fillId = GetOrCreateFillId(style.BackgroundColor);
        uint? nfId;
        if (style.NumberFormat == null)
        {
            nfId = null;
        }
        else
        {
            nfId = GetOrCreateNumberFormatId(style.NumberFormat);
        }

        var key = new CellFormatKey(
            fontId,
            fillId,
            nfId,
            style.Alignment.Horizontal,
            style.Alignment.Vertical,
            style.Alignment.WrapText);

        if (cellFormatIndex.TryGetValue(key, out var index))
        {
            return index;
        }

        index = (uint)cellFormats.Count;
        cellFormats.Add(key);
        cellFormatIndex[key] = index;
        return index;
    }

    uint GetOrCreateFontId(CellFont font)
    {
        var key = new FontKey(font.Bold, font.Underline, font.Color, font.Size, font.Name);
        if (fontIndex.TryGetValue(key, out var id))
        {
            return id;
        }

        id = (uint)fonts.Count;
        fonts.Add(key);
        fontIndex[key] = id;
        return id;
    }

    uint GetOrCreateFillId(string? backgroundColor)
    {
        if (backgroundColor == null)
        {
            return 0;
        }

        if (fillIndex.TryGetValue(backgroundColor, out var id))
        {
            return id;
        }

        id = (uint)fills.Count;
        fills.Add(backgroundColor);
        fillIndex[backgroundColor] = id;
        return id;
    }

    uint GetOrCreateNumberFormatId(string format)
    {
        if (numberFormatIds.TryGetValue(format, out var id))
        {
            return id;
        }

        id = nextNumberFormatId++;
        numberFormatIds[format] = id;
        customNumberFormats.Add(format);
        return id;
    }

    internal void WriteStylesXml(XmlWriter writer)
    {
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        writer.WriteStartElement("styleSheet", ns);

        // Number formats
        if (customNumberFormats.Count > 0)
        {
            writer.WriteStartElement("numFmts");
            writer.WriteAttributeString("count", customNumberFormats.Count.ToString());
            uint nfId = 164;
            foreach (var fmt in customNumberFormats)
            {
                writer.WriteStartElement("numFmt");
                writer.WriteAttributeString("numFmtId", nfId.ToString());
                writer.WriteAttributeString("formatCode", fmt);
                writer.WriteEndElement();
                nfId++;
            }
            writer.WriteEndElement();
        }

        // Fonts
        writer.WriteStartElement("fonts");
        writer.WriteAttributeString("count", fonts.Count.ToString());
        foreach (var fontKey in fonts)
        {
            writer.WriteStartElement("font");
            if (fontKey.Bold)
            {
                writer.WriteStartElement("b");
                writer.WriteEndElement();
            }
            if (fontKey.Underline)
            {
                writer.WriteStartElement("u");
                writer.WriteEndElement();
            }
            writer.WriteStartElement("sz");
            writer.WriteAttributeString("val", (fontKey.Size ?? 11).ToString(System.Globalization.CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteStartElement("name");
            writer.WriteAttributeString("val", fontKey.Name ?? "Calibri");
            writer.WriteEndElement();
            if (fontKey.Color != null)
            {
                writer.WriteStartElement("color");
                writer.WriteAttributeString("rgb", fontKey.Color);
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // font
        }
        writer.WriteEndElement(); // fonts

        // Fills
        writer.WriteStartElement("fills");
        writer.WriteAttributeString("count", fills.Count.ToString());
        // First fill: none
        writer.WriteStartElement("fill");
        writer.WriteStartElement("patternFill");
        writer.WriteAttributeString("patternType", "none");
        writer.WriteEndElement();
        writer.WriteEndElement();
        // Second fill: gray125
        writer.WriteStartElement("fill");
        writer.WriteStartElement("patternFill");
        writer.WriteAttributeString("patternType", "gray125");
        writer.WriteEndElement();
        writer.WriteEndElement();
        // Custom fills
        for (var i = 2; i < fills.Count; i++)
        {
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "solid");
            writer.WriteStartElement("fgColor");
            writer.WriteAttributeString("rgb", fills[i]!);
            writer.WriteEndElement();
            writer.WriteEndElement(); // patternFill
            writer.WriteEndElement(); // fill
        }
        writer.WriteEndElement(); // fills

        // Borders (minimal - one empty border)
        writer.WriteStartElement("borders");
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("border");
        writer.WriteElementString("left", "");
        writer.WriteElementString("right", "");
        writer.WriteElementString("top", "");
        writer.WriteElementString("bottom", "");
        writer.WriteElementString("diagonal", "");
        writer.WriteEndElement(); // border
        writer.WriteEndElement(); // borders

        // Cell formats
        writer.WriteStartElement("cellXfs");
        writer.WriteAttributeString("count", cellFormats.Count.ToString());
        foreach (var key in cellFormats)
        {
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("fontId", key.FontId.ToString());
            writer.WriteAttributeString("fillId", key.FillId.ToString());
            writer.WriteAttributeString("borderId", "0");
            if (key.FontId > 0)
            {
                writer.WriteAttributeString("applyFont", "1");
            }
            if (key.FillId > 0)
            {
                writer.WriteAttributeString("applyFill", "1");
            }
            if (key.NumberFormatId.HasValue)
            {
                writer.WriteAttributeString("numFmtId", key.NumberFormatId.Value.ToString());
                writer.WriteAttributeString("applyNumberFormat", "1");
            }
            if (key.HAlign != HorizontalAlignment.General ||
                key.VAlign != VerticalAlignment.Bottom ||
                key.WrapText)
            {
                writer.WriteAttributeString("applyAlignment", "1");
                writer.WriteStartElement("alignment");
                writer.WriteAttributeString("horizontal", key.HAlign.ToString().ToLowerInvariant());
                writer.WriteAttributeString("vertical", key.VAlign.ToString().ToLowerInvariant());
                if (key.WrapText)
                {
                    writer.WriteAttributeString("wrapText", "1");
                }
                writer.WriteEndElement(); // alignment
            }
            writer.WriteEndElement(); // xf
        }
        writer.WriteEndElement(); // cellXfs

        writer.WriteEndElement(); // styleSheet
    }
}
