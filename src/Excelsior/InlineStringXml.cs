using System.Security;
using System.Xml;

namespace Excelsior;

/// <summary>
/// Builds XML fragments for InlineString cell content.
/// These fragments are the inner content of an &lt;is&gt; element.
/// </summary>
static class InlineStringXml
{
    internal static string SimpleText(string text) =>
        $"""<t xml:space="preserve">{SecurityElement.Escape(text)}</t>""";

    internal static string BulletList(IReadOnlyList<string> items)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                AppendRun(sb, "\n");
            }

            AppendRun(sb, "● ", bold: true);
            AppendRun(sb, items[i]);
        }

        return sb.ToString();
    }

    internal static string LinkList(IReadOnlyList<string> items)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < items.Count; i++)
        {
            if (i > 0)
            {
                AppendRun(sb, "\n");
            }

            AppendRun(sb, "● ", bold: true);
            AppendRun(sb, items[i], underline: true, color: "0563C1");
        }

        return sb.ToString();
    }

    internal static void AppendRun(StringBuilder sb, string text, bool bold = false, bool underline = false, string? color = null)
    {
        sb.Append("<r>");
        if (bold || underline || color != null)
        {
            sb.Append("<rPr>");
            if (bold)
            {
                sb.Append("<b/>");
            }

            if (underline)
            {
                sb.Append("<u/>");
            }

            if (color != null)
            {
                sb.Append($"""<color rgb="{color}"/>""");
            }

            sb.Append("</rPr>");
        }

        sb.Append($"""<t xml:space="preserve">{SecurityElement.Escape(text)}</t></r>""");
    }
}
