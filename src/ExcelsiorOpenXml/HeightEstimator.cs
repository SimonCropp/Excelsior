namespace ExcelsiorOpenXml;

static class HeightEstimator
{
    const double DefaultRowHeight = 15; // Excel default row height in points
    const double MaxRowHeight = 409; // Excel maximum row height

    /// <summary>
    /// Estimates row height based on content and column width.
    /// </summary>
    public static double EstimateHeight(string text, double columnWidth, bool wrapText)
    {
        if (!wrapText || string.IsNullOrEmpty(text))
            return DefaultRowHeight;

        // Count explicit newlines
        var explicitLines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries).Length;

        // Estimate wrapped lines based on column width
        // Rough approximation: each character is about 1.2 width units
        var charsPerLine = (int)(columnWidth / 1.2);
        if (charsPerLine < 1) charsPerLine = 1;

        var totalChars = text.Length;
        var wrappedLines = (totalChars + charsPerLine - 1) / charsPerLine;

        var totalLines = Math.Max(explicitLines, wrappedLines);

        // Each line is approximately 15 points
        var estimatedHeight = totalLines * DefaultRowHeight;

        // Cap at maximum
        return Math.Min(estimatedHeight, MaxRowHeight);
    }

    /// <summary>
    /// Estimates height for a cell value.
    /// </summary>
    public static double EstimateHeight(object? value, double columnWidth, bool wrapText)
    {
        if (value == null)
            return DefaultRowHeight;

        if (value is string str)
            return EstimateHeight(str, columnWidth, wrapText);

        // Non-string values are single line
        return DefaultRowHeight;
    }
}
