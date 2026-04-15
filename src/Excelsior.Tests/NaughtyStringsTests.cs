[TestFixture]
public class NaughtyStringsTests
{
    record Row(string Value);

    [Test]
    public async Task Test()
    {
        var rows = TheNaughtyStrings.All
            .Select(_ => new Row(_))
            .ToList();

        var builder = new BookBuilder();
        builder.AddSheet(rows);

        var book = await builder.Build();

        await Verify(book);
    }
}
