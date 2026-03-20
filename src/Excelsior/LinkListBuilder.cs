static class LinkListBuilder
{
    public static string Build(List<Link> links)
    {
        var builder = new StringBuilder();
        var first = true;
        foreach (var link in links)
        {
            if (!first)
            {
                builder.Append('\n');
            }

            first = false;
            builder.Append("● ");
            if (link.Text == link.Url)
            {
                builder.Append(link.Url);
            }
            else
            {
                builder.Append(link.Text);
                builder.Append(" (");
                builder.Append(link.Url);
                builder.Append(')');
            }
        }

        return builder.ToString();
    }
}
