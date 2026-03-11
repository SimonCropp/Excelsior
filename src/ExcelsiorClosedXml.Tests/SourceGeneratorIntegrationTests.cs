[TestFixture]
public class SourceGeneratorIntegrationTests
{
    [Test]
    public async Task GeneratedExtensionMethods()
    {
        #region SourceGeneratedUsage

        var builder = new BookBuilder();

        List<GeneratedTestModel> data =
        [
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 25 },
        ];

        var sheet = builder.AddSheet(data);
        sheet.NameColumn(_ => _.Heading = "Full Name");
        sheet.AgeOrder(1);
        sheet.NameOrder(2);
        sheet.AgeWidth(15);

        #endregion

        using var book = await builder.Build();

        await Verify(book);
    }
}

#region SourceGeneratedModel

[SheetModel]
public class GeneratedTestModel
{
    public required string Name { get; init; }
    public required int Age { get; init; }
}

#endregion
