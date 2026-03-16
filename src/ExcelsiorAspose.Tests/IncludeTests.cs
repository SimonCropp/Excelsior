// ReSharper disable NotAccessedPositionalProperty.Local
[TestFixture]
public class IncludeTests
{
    record Target(string Name, int Age, string Email);

    // ReSharper disable once NotAccessedPositionalProperty.Local
    record AttributeIncludeFalseTarget(string Name, [property: Column(Include = false)] int Age, string Email);

    static List<Target> Data() =>
    [
        new("Alice", 30, "alice@test.com"),
        new("Bob", 25, "bob@test.com")
    ];

    [Test]
    public async Task AllIncludedByDefault()
    {
        var builder = new BookBuilder();
        builder.AddSheet(Data());

        var book = await builder.Build();
        await Verify(book);
    }

    [Test]
    public async Task ExcludeOne()
    {
        #region IncludeExcludeOne

        var builder = new BookBuilder();
        var sheet = builder.AddSheet(Data());
        sheet.Include(_ => _.Age, false);

        #endregion

        var book = await builder.Build();
        await Verify(book);
    }

    [Test]
    public async Task ExcludeOneViaColumn()
    {
        #region IncludeExcludeOneViaColumn

        var builder = new BookBuilder();
        var sheet = builder.AddSheet(Data());
        sheet.Column(
            _ => _.Age,
            _ => _.Include = false);

        #endregion

        var book = await builder.Build();
        await Verify(book);
    }

    [Test]
    public async Task AttributeIncludeFalse()
    {
        var builder = new BookBuilder();
        builder.AddSheet<AttributeIncludeFalseTarget>(
        [
            new("Alice", 30, "alice@test.com"),
            new("Bob", 25, "bob@test.com")
        ]);

        var book = await builder.Build();
        await Verify(book);
    }
}
