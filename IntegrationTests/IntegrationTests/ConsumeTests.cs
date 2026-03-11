using Excelsior;
using ExcelsiorClosedXml;

[TestFixture]
public class ConsumeTests
{
    [Test]
    public async Task GeneratedExtensionMethods()
    {
        var builder = new BookBuilder();

        List<TestModel> data =
        [
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 25 },
        ];

        var sheet = builder.AddSheet(data);
        sheet.NameColumn(c => c.Heading = "Full Name");
        sheet.AgeOrder(1);
        sheet.NameOrder(2);
        sheet.AgeWidth(15);

        using var book = await builder.Build();
    }
}

[SheetModel]
public class TestModel
{
    public required string Name { get; init; }
    public required int Age { get; init; }
}
