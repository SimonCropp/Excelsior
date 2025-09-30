[TestFixture]
public class NullableTests
{
    [Test]
    public async Task Nulls()
    {
        var builder = new BookBuilder();
        builder.AddSheet(NullableTargets.Data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WithRender()
    {
        var builder = new BookBuilder();
        var sheet = builder.AddSheet(NullableTargets.Data);
        sheet.Column(
            _ => _.Number,
            _ => _.Render = (_, value) => value.ToString());
        sheet.Column(
            _ => _.DateTime,
            _ => _.Render = (_, value) => value?.ToString(CultureInfo.InvariantCulture));
        sheet.Column(
            _ => _.Enum,
            _ => _.Render = (_, value) => value.ToString());
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
        builder.AddSheet(NullableTargets.Data)
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
        var sheetBuilder = bookBuilder.AddSheet(NullableTargets.Data);
        sheetBuilder.NullDisplay(_ => _.Number, "[No Number]");
        sheetBuilder.NullDisplay(_ => _.String, "[No String]");
        sheetBuilder.NullDisplay(_ => _.DateTime, "[No DateTime]");
        sheetBuilder.NullDisplay(_ => _.Enum, "[No Enum]");
        sheetBuilder.NullDisplay(_ => _.Bool, "[No Bool]");

        var book = await bookBuilder.Build();

        await Verify(book);
    }
}