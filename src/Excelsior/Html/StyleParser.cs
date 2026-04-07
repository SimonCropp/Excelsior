namespace Excelsior;

static class StyleParser
{
    internal static Dictionary<string, string> Parse(string? style)
    {
        var result = new Dictionary<string, string>(8, StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style))
        {
            return result;
        }

        var remaining = style.AsSpan();
        while (remaining.Length > 0)
        {
            var semiIndex = remaining.IndexOf(';');
            var declaration = semiIndex >= 0 ? remaining[..semiIndex] : remaining;
            remaining = semiIndex >= 0 ? remaining[(semiIndex + 1)..] : default;

            declaration = declaration.Trim();
            if (declaration.IsEmpty)
            {
                continue;
            }

            var colonIndex = declaration.IndexOf(':');
            if (colonIndex <= 0)
            {
                continue;
            }

            var property = declaration[..colonIndex].Trim();
            var value = declaration[(colonIndex + 1)..].Trim();
            if (!property.IsEmpty && !value.IsEmpty)
            {
                result[property.ToString()] = value.ToString();
            }
        }

        return result;
    }

    internal static double? ParseFontSize(string value)
    {
        var span = value.AsSpan().Trim();

        if (TryParseSuffix(span, "pt", out var pt))
        {
            return pt;
        }

        if (TryParseSuffix(span, "px", out var px))
        {
            return px * 0.75;
        }

        if (TryParseSuffix(span, "em", out var em))
        {
            return em * 12;
        }

        if (double.TryParse(span.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var raw))
        {
            return raw;
        }

        if (span.Equals("xx-small", StringComparison.OrdinalIgnoreCase))
        {
            return 7;
        }

        if (span.Equals("x-small", StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }

        if (span.Equals("small", StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }

        if (span.Equals("medium", StringComparison.OrdinalIgnoreCase))
        {
            return 12;
        }

        if (span.Equals("large", StringComparison.OrdinalIgnoreCase))
        {
            return 14;
        }

        if (span.Equals("x-large", StringComparison.OrdinalIgnoreCase))
        {
            return 18;
        }

        if (span.Equals("xx-large", StringComparison.OrdinalIgnoreCase))
        {
            return 24;
        }

        return null;
    }

    static bool TryParseSuffix(ReadOnlySpan<char> span, string suffix, out double result)
    {
        if (span.EndsWith(suffix.AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            return double.TryParse(
                span[..^suffix.Length].Trim().ToString(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }

        result = 0;
        return false;
    }
}
