static class RichText
{
    public static void Enumerable(Range cell, IEnumerable<string> enumerable, bool trimWhitespace)
    {
        var boldFont = cell.Worksheet.Workbook.CreateFont();
        boldFont.Bold = true;
        boldFont.Size = 11;
        var rich = cell.RichText;
        var list = enumerable.ToList();
        var builder = new StringBuilder();
        var normalFont = cell.Worksheet.Workbook.CreateFont();
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            rich.Append("• ", boldFont);
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

            rich.Append(builder.ToString(), normalFont);
            builder.Clear();
        }
    }
}