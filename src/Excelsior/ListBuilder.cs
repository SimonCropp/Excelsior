static class ListBuilder
{
    public static string Build(IEnumerable<string> enumerable, bool trimWhitespace)
    {
        if (trimWhitespace)
        {
            enumerable = enumerable.Select(_ => _.Trim());
        }

        return string.Join('\n', enumerable.Select(_ => $"● {_}"));
    }
}