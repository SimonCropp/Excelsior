using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace Excelsior;

static class HtmlSegmentParser
{
    static readonly HtmlParser parser = new();

    internal static List<TextSegment> Parse(string html)
    {
        var document = parser.ParseDocument(string.Concat("<body>", html, "</body>"));
        var body = document.Body!;
        var segments = new List<TextSegment>();
        var format = new FormatState();
        ProcessNode(body, format, segments, false);
        TrimTrailingNewlines(segments);
        return segments;
    }

    static void ProcessNode(INode node, FormatState format, List<TextSegment> segments, bool inPre)
    {
        foreach (var child in node.ChildNodes)
        {
            switch (child)
            {
                case IText textNode:
                {
                    var text = inPre ? textNode.Data : CollapseWhitespace(textNode.Data);
                    if (text.Length > 0 &&
                        !(string.IsNullOrWhiteSpace(text) &&
                          IsInterBlockWhitespace(textNode)))
                    {
                        segments.Add(new(text, format));
                    }

                    break;
                }
                case IElement element:
                    ProcessElement(element, format, segments, inPre);
                    break;
            }
        }
    }

    static void ProcessElement(IElement element, FormatState format, List<TextSegment> segments, bool inPre)
    {
        var tag = element.LocalName;
        var newFormat = format;
        ApplyElementFormatting(element, tag, ref newFormat);

        switch (tag)
        {
            case "br":
                segments.Add(new("\n", format));
                return;
            case "hr":
                EnsureNewline(segments, format);
                segments.Add(new("———", format));
                segments.Add(new("\n", format));
                return;
            case "img":
            {
                var alt = element.GetAttribute("alt");
                if (!string.IsNullOrEmpty(alt))
                {
                    segments.Add(new(alt!, format));
                }

                return;
            }
            case "svg":
                return;
        }

        var isBlock = IsBlockElement(tag);
        if (isBlock)
        {
            EnsureNewline(segments, format);
        }

        switch (tag)
        {
            case "li":
            {
                var parent = element.ParentElement?.LocalName;
                var depth = GetListDepth(element);
                var indent = depth > 0 ? new string(' ', depth * 2) : "";
                if (parent == "ol")
                {
                    var index = 1;
                    foreach (var sibling in element.ParentElement!.Children)
                    {
                        if (sibling == element)
                        {
                            break;
                        }

                        if (sibling.LocalName == "li")
                        {
                            index++;
                        }
                    }

                    segments.Add(new($"{indent}{index}. ", newFormat));
                }
                else
                {
                    var bullet = depth switch
                    {
                        0 => "●",
                        1 => "○",
                        _ => "■"
                    };
                    segments.Add(new($"{indent}{bullet} ", newFormat));
                }

                ProcessNode(element, newFormat, segments, inPre);
                EnsureNewline(segments, format);
                return;
            }
            case "a":
            {
                ProcessNode(element, newFormat, segments, inPre);
                var href = element.GetAttribute("href");
                if (!string.IsNullOrEmpty(href))
                {
                    var linkText = element.TextContent.Trim();
                    if (linkText != href)
                    {
                        segments.Add(new($" ({href})", format));
                    }
                }

                return;
            }
            case "q":
            {
                segments.Add(new("\u201C", format));
                ProcessNode(element, newFormat, segments, inPre);
                segments.Add(new("\u201D", format));
                return;
            }
            case "td" or "th":
            {
                ProcessNode(element, newFormat, segments, inPre);
                var nextSibling = element.NextElementSibling;
                if (nextSibling is { LocalName: "td" or "th" })
                {
                    segments.Add(new("\t", format));
                }

                return;
            }
            case "tr":
            {
                ProcessNode(element, newFormat, segments, inPre);
                EnsureNewline(segments, format);
                return;
            }
            case "col":
                return;
            case "pre":
                ProcessNode(element, newFormat, segments, true);
                break;
            default:
                ProcessNode(element, newFormat, segments, inPre);
                break;
        }

        if (isBlock)
        {
            EnsureNewline(segments, format);
        }
    }

    static void ApplyElementFormatting(IElement element, string tag, ref FormatState format)
    {
        switch (tag)
        {
            case "b" or "strong" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "th" or "caption" or "dt":
                format.Bold = true;
                break;
            case "i" or "em" or "cite" or "dfn" or "var":
                format.Italic = true;
                break;
            case "u" or "ins":
                format.Underline = true;
                break;
            case "s" or "strike" or "del":
                format.Strikethrough = true;
                break;
            case "a":
                format.Underline = true;
                format.Color ??= "0563C1";
                format.LinkUrl = element.GetAttribute("href");
                break;
            case "sup":
                format.Superscript = true;
                format.Subscript = false;
                break;
            case "sub":
                format.Subscript = true;
                format.Superscript = false;
                break;
            case "mark":
                format.Color ??= "000000";
                break;
            case "code" or "kbd" or "samp":
                format.FontFamily ??= "Courier New";
                break;
            case "small":
                format.FontSizePt = Math.Round((format.FontSizePt ?? 12) * 0.8, 2);
                break;
        }

        var fontColor = element.GetAttribute("color");
        if (fontColor != null)
        {
            var parsed = ColorParser.Parse(fontColor);
            if (parsed != null)
            {
                format.Color = parsed;
            }
        }

        var fontFace = element.GetAttribute("face");
        if (fontFace != null)
        {
            format.FontFamily = fontFace;
        }

        var fontSize = element.GetAttribute("size");
        if (fontSize != null && double.TryParse(fontSize, CultureInfo.InvariantCulture, out var size))
        {
            format.FontSizePt = size;
        }

        ApplyInlineStyle(element, ref format);
    }

    static void ApplyInlineStyle(IElement element, ref FormatState format)
    {
        var style = element.GetAttribute("style");
        if (style == null)
        {
            return;
        }

        var declarations = StyleParser.Parse(style);

        if (declarations.TryGetValue("font-weight", out var fontWeight))
        {
            if (fontWeight is "bold" or "bolder" or "700" or "800" or "900")
            {
                format.Bold = true;
            }
        }

        if (declarations.TryGetValue("font-style", out var fontStyle))
        {
            if (fontStyle is "italic" or "oblique")
            {
                format.Italic = true;
            }
        }

        if (declarations.TryGetValue("text-decoration", out var textDecoration))
        {
            if (textDecoration.Contains("underline", StringComparison.OrdinalIgnoreCase))
            {
                format.Underline = true;
            }

            if (textDecoration.Contains("line-through", StringComparison.OrdinalIgnoreCase))
            {
                format.Strikethrough = true;
            }
        }

        if (declarations.TryGetValue("color", out var color))
        {
            var parsed = ColorParser.Parse(color);
            if (parsed != null)
            {
                format.Color = parsed;
            }
        }

        if (declarations.TryGetValue("font-size", out var fontSizeStr))
        {
            var parsed = StyleParser.ParseFontSize(fontSizeStr);
            if (parsed != null)
            {
                format.FontSizePt = parsed;
            }
        }

        if (declarations.TryGetValue("font-family", out var fontFamily))
        {
            format.FontFamily = fontFamily.Trim('\'', '"');
        }

        if (declarations.TryGetValue("vertical-align", out var verticalAlign))
        {
            switch (verticalAlign)
            {
                case "super":
                    format.Superscript = true;
                    format.Subscript = false;
                    break;
                case "sub":
                    format.Subscript = true;
                    format.Superscript = false;
                    break;
            }
        }
    }

    static bool IsBlockElement(string tag) =>
        tag is "p" or "div" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6"
            or "ul" or "ol" or "li" or "blockquote" or "pre" or "table"
            or "tr" or "hr" or "section" or "article" or "header" or "footer"
            or "nav" or "aside" or "main" or "figure" or "figcaption" or "details"
            or "summary" or "address" or "dt" or "dd" or "dl"
            or "caption" or "tbody" or "thead" or "tfoot"
            or "body" or "html";

    static bool IsInterBlockWhitespace(IText textNode)
    {
        var parent = textNode.ParentElement;
        if (parent == null)
        {
            return false;
        }

        var hasChildren = false;
        foreach (var child in parent.Children)
        {
            hasChildren = true;
            if (!IsBlockElement(child.LocalName))
            {
                return false;
            }
        }

        return hasChildren;
    }

    static string CollapseWhitespace(string text)
    {
        var builder = new StringBuilder(text.Length);
        var lastWasSpace = false;
        foreach (var c in text)
        {
            if (char.IsWhiteSpace(c))
            {
                if (!lastWasSpace)
                {
                    builder.Append(' ');
                    lastWasSpace = true;
                }
            }
            else
            {
                builder.Append(c);
                lastWasSpace = false;
            }
        }

        return builder.ToString();
    }

    static void EnsureNewline(List<TextSegment> segments, FormatState format)
    {
        if (segments.Count == 0)
        {
            return;
        }

        var last = segments[^1];
        if (!last.Text.EndsWith('\n'))
        {
            segments.Add(new("\n", format));
        }
    }

    static int GetListDepth(IElement listItem)
    {
        var depth = 0;
        var current = listItem.ParentElement;
        while (current != null)
        {
            if (current.LocalName is "ul" or "ol")
            {
                depth++;
            }

            current = current.ParentElement;
        }

        // Subtract 1 because the immediate parent ul/ol is depth 0
        return Math.Max(0, depth - 1);
    }

    static void TrimTrailingNewlines(List<TextSegment> segments)
    {
        while (segments.Count > 0 && segments[^1].Text == "\n")
        {
            segments.RemoveAt(segments.Count - 1);
        }
    }
}
