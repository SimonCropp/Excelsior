static class CellWrite
{
    public static void String(Cell cell, string value)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new(BuildText(value));
    }

    public static void Html(Cell cell, string value) =>
        SpreadsheetHtmlConverter.SetCellHtml(cell, value);

    public static void StringOrHtml(Cell cell, string value, bool isHtml)
    {
        if (isHtml)
        {
            Html(cell, value);
        }
        else
        {
            String(cell, value);
        }
    }

    public static Text BuildText(string value) =>
        new(StripInvalidXmlChars(NormalizeNewlines(value)))
        {
            Space = SpaceProcessingModeValues.Preserve
        };

    public static string NormalizeNewlines(string value)
    {
        if (value.AsSpan().IndexOf('\r') < 0)
        {
            return value;
        }

        return value.Replace("\r\n", "\n").Replace('\r', '\n');
    }

    // Strips characters that XML 1.0 forbids (most C0 controls, lone surrogates,
    // 0xFFFE/0xFFFF). Without this, any such char in a cell value crashes
    // OpenXml's serializer with InvalidXmlChar at save time.
    // TODO: revert once https://github.com/dotnet/Open-XML-SDK/issues/1532 ships —
    // OpenXml itself should escape these via the OOXML _xHHHH_ convention.
    public static string StripInvalidXmlChars(string value)
    {
        if (value.Length == 0)
        {
            return value;
        }

        StringBuilder? builder = null;
        var i = 0;
        while (i < value.Length)
        {
            var c = value[i];
            int advance;
            bool valid;

            if (char.IsHighSurrogate(c) &&
                i + 1 < value.Length &&
                char.IsLowSurrogate(value[i + 1]))
            {
                advance = 2;
                valid = true;
            }
            else if (char.IsSurrogate(c))
            {
                advance = 1;
                valid = false;
            }
            else
            {
                advance = 1;
                valid = c == '\t' ||
                        c == '\n' ||
                        c == '\r' ||
                        (c >= 0x20 && c <= 0xD7FF) ||
                        (c >= 0xE000 && c <= 0xFFFD);
            }

            if (valid)
            {
                builder?.Append(value, i, advance);
            }
            else if (builder == null)
            {
                builder = new(value.Length);
                builder.Append(value, 0, i);
            }

            i += advance;
        }

        return builder?.ToString() ?? value;
    }
}
