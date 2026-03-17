static class ListBuilder
{
    public static string Build(IEnumerable<string?> enumerable)
    {
        var trim = ValueRenderer.TrimWhitespace;
        var builder = new StringBuilder();
        var first = true;
        foreach (var item in enumerable)
        {
            if (!first)
            {
                builder.Append('\n');
            }

            first = false;
            builder.Append("● ");
            builder.Append(trim ? item?.Trim() : item);
        }

        return builder.ToString();
    }
}