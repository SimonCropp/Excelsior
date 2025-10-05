static class RichText
{
    public static void Enumerable(Cell cell, IEnumerable<string> enumerable, bool trimWhitespace)
    {
        var rich = cell.CreateRichText();
        var list = enumerable.ToList();
        var builder = new StringBuilder();
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            rich.AddText("• ").SetBold();
            foreach (var line in item.AsSpan().EnumerateLines())
            {
                if (trimWhitespace)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    builder.Append(line.Trim());
                }
                else
                {
                    builder.Append(line);
                }

                builder.Append("\n   ");
            }

            if (index < list.Count - 1)
            {
                builder.Length -= 3;
            }
            else
            {
                builder.Length -= 4;
            }

            rich.AddText(builder.ToString());
            builder.Clear();
        }
    }
}