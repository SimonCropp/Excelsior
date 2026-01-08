namespace ExcelsiorOpenXml;

static class WidthEstimator
{
    /// <summary>
    /// Estimates column width based on text content.
    /// Excel column width units are approximately 1/256th of the width of the '0' character
    /// in the default font. For practical purposes, we use character count * 1.2 as a rough estimate.
    /// </summary>
    public static double EstimateWidth(string text)
    {
        if (string.IsNullOrEmpty(text))
            return 8; // Minimum width

        // Split by newlines and find the longest line
        var lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        var maxLength = lines.Length > 0 ? lines.Max(l => l.Length) : text.Length;

        // Character width multiplier (approximate)
        // Excel units are complex, but roughly 1 char ≈ 1.2 width units
        var estimatedWidth = maxLength * 1.2;

        // Add some padding
        return estimatedWidth + 2;
    }

    /// <summary>
    /// Estimates width for a cell value (handles different types).
    /// </summary>
    public static double EstimateWidth(object? value)
    {
        if (value == null)
            return 8;

        return value switch
        {
            string str => EstimateWidth(str),
            DateTime => 20, // Date/time format is typically around 20 chars
            DateTimeOffset => 25,
            bool => 8,
            int or long or short or byte => 12,
            decimal or double or float => 15,
            _ => EstimateWidth(value.ToString() ?? "")
        };
    }
}
