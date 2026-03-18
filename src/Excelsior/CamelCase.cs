static class CamelCase
{
    public static string Split(string text)
    {
        var spaceCount = 0;
        for (var i = 1; i < text.Length; i++)
        {
            if (char.IsUpper(text[i]))
            {
                spaceCount++;
            }
        }

        if (spaceCount == 0)
        {
            return text;
        }

        var result = new char[text.Length + spaceCount];
        var pos = 0;
        result[pos++] = text[0];
        for (var i = 1; i < text.Length; i++)
        {
            if (char.IsUpper(text[i]))
            {
                result[pos++] = ' ';
            }

            result[pos++] = text[i];
        }

        return new(result);
    }
}
