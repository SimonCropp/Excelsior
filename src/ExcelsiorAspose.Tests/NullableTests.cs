[TestFixture]
public class NullableTests
{
    [Test]
    public async Task Nulls()
    {
        var builder = new BookBuilder();
        builder.AddSheet(Data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WithRender()
    {
        var builder = new BookBuilder();
        var sheet = builder.AddSheet(Data);
        sheet.Column(
            _ => _.Number,
            _ => _.Render = (_, number) => number.ToString());
        sheet.Column(
            _ => _.DateTime,
            _ => _.Render = (_, dateTime) => dateTime?.ToString(CultureInfo.InvariantCulture));
        sheet.Column(
            _ => _.Enum,
            _ => _.Render = (_, enumValue) => enumValue.ToString());
        sheet.Column(
            _ => _.String,
            _ => _.Render = (_, value) => value?.ToString());
        sheet.Column(
            _ => _.Bool,
            _ => _.Render = (_, value) => value?.ToString().ToUpper());
        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullDisplay()
    {
        var builder = new BookBuilder();
        builder.AddSheet(Data)
            .Column(_ => _.Number, _ => _.NullDisplay = "[No Number]")
            .Column(_ => _.String, _ => _.NullDisplay = "[No String]")
            .Column(_ => _.DateTime, _ => _.NullDisplay = "[No DateTime]")
            .Column(_ => _.Enum, _ => _.NullDisplay = "[No Enum]")
            .Column(_ => _.Bool, _ => _.NullDisplay = "[No Bool]");

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullDisplayShortcuts()
    {
        var bookBuilder = new BookBuilder();
        var sheetBuilder = bookBuilder.AddSheet(Data);
        sheetBuilder.NullDisplay(_ => _.Number, "[No Number]");
        sheetBuilder.NullDisplay(_ => _.String, "[No String]");
        sheetBuilder.NullDisplay(_ => _.DateTime, "[No DateTime]");
        sheetBuilder.NullDisplay(_ => _.Enum, "[No Enum]");
        sheetBuilder.NullDisplay(_ => _.Bool, "[No Bool]");

        var book = await bookBuilder.Build();

        await Verify(book);
    }

    static IReadOnlyList<NullableTargets> Data { get; } =
    [
        new()
        {
            Number = null,
            String = null,
            DateTime = null,
            Enum = null,
            Bool = null
        },
        new()
        {
            Number = 1,
            String = "value",
            DateTime = new DateTime(2020, 1, 1),
            Enum = AnEnum.Value,
            Bool = true
        },
    ];

    class NullableTargets
    {
        public required int? Number { get; init; }
        public required string? String { get; init; }
        public required DateTime? DateTime { get; init; }
        public required AnEnum? Enum { get; init; }
        public required bool? Bool { get; init; }

    }

    enum AnEnum
    {
        Value
    }
}