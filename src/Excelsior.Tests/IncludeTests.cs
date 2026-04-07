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

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task ExcludeOne()
    {
        #region IncludeExcludeOne

        List<Target> data = [
            new("Alice", 30, "alice@test.com"),
            new("Bob", 25, "bob@test.com")
        ];
        var builder = new BookBuilder();
        var sheet = builder.AddSheet(data);
        sheet.Exclude(_ => _.Age);

        #endregion

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task ExcludeOneViaColumn()
    {
        #region IncludeExcludeOneViaColumn

        List<Target> data = [
            new("Alice", 30, "alice@test.com"),
            new("Bob", 25, "bob@test.com")
        ];
        var builder = new BookBuilder();
        var sheet = builder.AddSheet(data);
        sheet.Column(
            _ => _.Age,
            _ => _.Include = false);

        #endregion

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
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

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task ToggleBasedOnState()
    {
        #region IncludeToggleBasedOnState

        var data = Data();
        var isInternalReport = true;

        var builder = new BookBuilder();
        var sheet = builder.AddSheet(data);
        sheet.Include(_ => _.Email, !isInternalReport);

        #endregion

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task MultipleSpreadsheetsSameModel_Public()
    {
        #region IncludeMultipleSpreadsheets_Public

        var data = Data();

        // Public report: exclude age and email
        var builder = new BookBuilder();
        var sheet = builder.AddSheet(data);
        sheet.Exclude(_ => _.Age);
        sheet.Exclude(_ => _.Email);

        #endregion

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task MultipleSpreadsheetsSameModel_Internal()
    {
        #region IncludeMultipleSpreadsheets_Internal

        List<Target> data = [
            new("Alice", 30, "alice@test.com"),
            new("Bob", 25, "bob@test.com")
        ];

        // Internal report: include all columns
        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        using var stream = await builder.Build();
        await Verify(stream, "xlsx");
    }
}
