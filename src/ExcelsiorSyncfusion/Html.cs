static class Html
{
    public static string Enumerable(IEnumerable<string> enumerable, bool trimWhitespace)
    {
        var list = enumerable.ToList();
        var builder = new StringBuilder(
            """
            <ul>

            """);
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            builder.Append("<li>");

            if (trimWhitespace)
            {
                item = item.Trim();
            }

            item = WebUtility.HtmlEncode(item);

            builder.Append(item);

            builder.AppendLine("</li>");
        }

        builder.Append("</ul>");

        return builder.ToString();
    }
}