using System.Security;

namespace Excelsior;

/// <summary>
/// Converts HTML content to InlineString XML for xlsx cells.
/// </summary>
static class HtmlToInlineString
{
    /// <summary>
    /// Converts HTML to InlineString XML fragment (inner content of &lt;is&gt; element).
    /// </summary>
    internal static string Convert(string html)
    {
        var segments = HtmlSegmentParser.Parse(html);
        var sb = new StringBuilder();

        foreach (var segment in segments)
        {
            sb.Append("<r>");
            if (segment.Format.HasFormatting)
            {
                AppendRunProperties(sb, segment.Format);
            }

            sb.Append($"""<t xml:space="preserve">{SecurityElement.Escape(segment.Text)}</t>""");
            sb.Append("</r>");
        }

        return sb.ToString();
    }

    static void AppendRunProperties(StringBuilder sb, FormatState format)
    {
        sb.Append("<rPr>");

        if (format.Bold)
        {
            sb.Append("<b/>");
        }

        if (format.Italic)
        {
            sb.Append("<i/>");
        }

        if (format.Underline)
        {
            sb.Append("<u/>");
        }

        if (format.Strikethrough)
        {
            sb.Append("<strike/>");
        }

        if (format.Color != null)
        {
            sb.Append($"""<color rgb="FF{format.Color}"/>""");
        }

        if (format.FontSizePt != null)
        {
            sb.Append(string.Create(CultureInfo.InvariantCulture, $"""<sz val="{format.FontSizePt.Value}"/>"""));
        }

        if (format.FontFamily != null)
        {
            sb.Append($"""<rFont val="{SecurityElement.Escape(format.FontFamily)}"/>""");
        }

        if (format.Superscript)
        {
            sb.Append("""<vertAlign val="superscript"/>""");
        }
        else if (format.Subscript)
        {
            sb.Append("""<vertAlign val="subscript"/>""");
        }

        sb.Append("</rPr>");
    }
}
