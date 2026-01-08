using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelsiorOpenXml;

static partial class HtmlConverter
{
    /// <summary>
    /// Converts basic HTML to OpenXML runs with formatting.
    /// Supports: &lt;b>, &lt;i>, &lt;u>, &lt;br>, &lt;font color="...">, &lt;strong>, &lt;em>
    /// </summary>
    public static List<Run> ConvertToRuns(string html)
    {
        var runs = new List<Run>();

        try
        {
            // Parse HTML and create runs
            ParseHtml(html, runs);
        }
        catch
        {
            // Fallback: create a single run with plain text
            runs.Clear();
            runs.Add(CreateRun(StripHtmlTags(html), new FormatState()));
        }

        return runs;
    }

    static void ParseHtml(string html, List<Run> runs)
    {
        var formatStack = new Stack<FormatState>();
        formatStack.Push(new FormatState());

        var position = 0;
        var textBuffer = new StringBuilder();

        while (position < html.Length)
        {
            var nextTag = FindNextTag(html, position);

            if (nextTag.Start == -1)
            {
                // No more tags, add remaining text
                textBuffer.Append(html[position..]);
                break;
            }

            // Add text before tag
            if (nextTag.Start > position)
            {
                textBuffer.Append(html[position..nextTag.Start]);
            }

            // Process tag
            if (nextTag.IsClosing)
            {
                // Flush text buffer with current format
                if (textBuffer.Length > 0)
                {
                    runs.Add(CreateRun(textBuffer.ToString(), formatStack.Peek()));
                    textBuffer.Clear();
                }

                // Pop format (but keep at least one)
                if (formatStack.Count > 1)
                    formatStack.Pop();
            }
            else if (nextTag.Tag == "br")
            {
                // Add line break
                if (textBuffer.Length > 0)
                {
                    runs.Add(CreateRun(textBuffer.ToString(), formatStack.Peek()));
                    textBuffer.Clear();
                }
                runs.Add(CreateLineBreak());
            }
            else
            {
                // Flush text buffer with current format before pushing new format
                if (textBuffer.Length > 0)
                {
                    runs.Add(CreateRun(textBuffer.ToString(), formatStack.Peek()));
                    textBuffer.Clear();
                }

                // Push new format
                var newFormat = formatStack.Peek().Clone();
                ApplyTag(newFormat, nextTag.Tag, nextTag.Attributes);
                formatStack.Push(newFormat);
            }

            position = nextTag.End;
        }

        // Flush remaining text
        if (textBuffer.Length > 0)
        {
            runs.Add(CreateRun(textBuffer.ToString(), formatStack.Peek()));
        }
    }

    static (int Start, int End, string Tag, bool IsClosing, Dictionary<string, string> Attributes) FindNextTag(string html, int startPos)
    {
        var tagStart = html.IndexOf('<', startPos);
        if (tagStart == -1)
            return (-1, -1, "", false, new());

        var tagEnd = html.IndexOf('>', tagStart);
        if (tagEnd == -1)
            return (-1, -1, "", false, new());

        var tagContent = html.Substring(tagStart + 1, tagEnd - tagStart - 1).Trim();
        var isClosing = tagContent.StartsWith('/');

        if (isClosing)
            tagContent = tagContent[1..].Trim();

        // Parse tag name and attributes
        var parts = tagContent.Split([' '], 2);
        var tagName = parts[0].ToLowerInvariant();
        var attributes = new Dictionary<string, string>();

        if (parts.Length > 1 && !isClosing)
        {
            // Simple attribute parsing (handles color="#rrggbb" and color="red")
            var attrMatch = ColorAttributeRegex().Match(parts[1]);
            if (attrMatch.Success)
            {
                attributes["color"] = attrMatch.Groups[1].Value;
            }
        }

        return (tagStart, tagEnd + 1, tagName, isClosing, attributes);
    }

    static void ApplyTag(FormatState format, string tag, Dictionary<string, string> attributes)
    {
        switch (tag)
        {
            case "b":
            case "strong":
                format.Bold = true;
                break;
            case "i":
            case "em":
                format.Italic = true;
                break;
            case "u":
                format.Underline = true;
                break;
            case "font":
                if (attributes.TryGetValue("color", out var color))
                {
                    format.Color = color;
                }
                break;
        }
    }

    static Run CreateRun(string text, FormatState format)
    {
        var run = new Run();

        var runProperties = new RunProperties();
        var hasProperties = false;

        if (format.Bold)
        {
            runProperties.Append(new Bold());
            hasProperties = true;
        }

        if (format.Italic)
        {
            runProperties.Append(new Italic());
            hasProperties = true;
        }

        if (format.Underline)
        {
            runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Underline());
            hasProperties = true;
        }

        if (format.Color != null)
        {
            var colorHex = ParseColor(format.Color);
            if (colorHex != null)
            {
                runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = colorHex });
                hasProperties = true;
            }
        }

        if (hasProperties)
            run.Append(runProperties);

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        return run;
    }

    static Run CreateLineBreak()
    {
        var run = new Run();
        run.Append(new Text("\n") { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    static string? ParseColor(string color)
    {
        // Handle hex colors like "#FF0000" or "FF0000"
        if (color.StartsWith('#'))
            color = color[1..];

        if (color.Length == 6)
        {
            return "FF" + color.ToUpperInvariant(); // Add alpha channel
        }

        // Handle named colors (basic set)
        return color.ToLowerInvariant() switch
        {
            "red" => "FFFF0000",
            "blue" => "FF0000FF",
            "green" => "FF00FF00",
            "yellow" => "FFFFFF00",
            "white" => "FFFFFFFF",
            "black" => "FF000000",
            _ => null
        };
    }

    static string StripHtmlTags(string html) =>
        HtmlTagRegex().Replace(html, "");

    [GeneratedRegex(@"color=[""']([^""']+)[""']", RegexOptions.IgnoreCase)]
    private static partial Regex ColorAttributeRegex();

    [GeneratedRegex(@"<[^>]+>")]
    private static partial Regex HtmlTagRegex();

    class FormatState
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string? Color { get; set; }

        public FormatState Clone() => new()
        {
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Color = Color
        };
    }
}
