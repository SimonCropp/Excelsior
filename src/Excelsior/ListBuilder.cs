static class ListBuilder
{
    public static string Build(IEnumerable<string> enumerable)
    {
        if (ValueRenderer.TrimWhitespace)
        {
            enumerable = enumerable.Select(_ => _.Trim());
        }

        return string.Join('\n', enumerable.Select(_ => $"● {_}"));
    }
}