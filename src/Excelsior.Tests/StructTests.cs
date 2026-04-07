[TestFixture]
public class StructTests
{
    [Test]
    public async Task Test()
    {
        var builder = new BookBuilder();

        List<Target> data =
        [
            new()
            {
                Name = "John Doe",
            },
        ];
        builder.AddSheet(data);

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    public readonly struct Target
    {
        public required string Name { get; init; }
    }
}