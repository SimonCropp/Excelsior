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
    [Test]
    public async Task ColumnAttributesAppliedAutomatically()
    {
        var builder = new BookBuilder();

        List<GeneratedColumnAttributeModel> data =
        [
            new()
            {
                Id = 1,
                Name = "Alice",
                Email = "alice@test.com",
                HireDate = new(2020, 1, 15)
            },
            new()
            {
                Id = 2,
                Name = "Bob",
                Email = "bob@test.com",
                HireDate = null
            },
        ];

        builder.AddSheet(data);

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

[SheetModel]
public class GeneratedColumnAttributeModel
{
    [Column(Heading = "Employee ID", Order = 1, Format = "0000")]
    public required int Id { get; init; }

    [Column(Heading = "Full Name", Order = 2, Width = 20)]
    public required string Name { get; init; }

    [Column(Heading = "Email Address", Width = 30)]
    public required string Email { get; init; }

    [Column(Order = 3, NullDisplay = "unknown")]
    public Date? HireDate { get; init; }
}
