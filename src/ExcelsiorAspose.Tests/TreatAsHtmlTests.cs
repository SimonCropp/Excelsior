[TestFixture]
public class TreatAsHtmlTests
{
    public record Target(string Value1, string Value2 = "sss");

    [Test]
    public async Task Test()
    {
        var bookBuilder = new BookBuilder();

        List<Target> data =
        [
            new("<ul><li>the value</li></ul>"),
        ];
        var sheetBuilder = bookBuilder.AddSheet(data);

        sheetBuilder.Column(
            _ => _.Value1,
            _ => _.TreatAsHtml = true);
        var book = await bookBuilder.Build();

        await Verify(book);
    }
    [Test]
    public async Task LongText()
    {
        var bookBuilder = new BookBuilder();

        List<Target> data =
        [
            new("aaaaaaaaaaaaaaaaaaaaaaaaaaa bbbbbbbbbbbbbbbbbbbbbbbb cccccccccccccccccccccccccccc"),
        ];
        var sheetBuilder = bookBuilder.AddSheet(data);

        sheetBuilder.Column(
            _ => _.Value1,
            _ => _.TreatAsHtml = true);
        var book = await bookBuilder.Build();

        await Verify(book);
    }

    [Test]
    public async Task LongTextNarrowColumn()
    {
        var bookBuilder = new BookBuilder();

        List<Target> data =
        [
            new("aaaaaaaaaaaaaaaaaaaaaaaaaaa bbbbbbbbbbbbbbbbbbbbbbbb cccccccccccccccccccccccccccc"),
        ];
        var sheetBuilder = bookBuilder.AddSheet(data);

        sheetBuilder.Column(
            _ => _.Value1,
            _ =>
            {
                _.Width = 20;
                _.TreatAsHtml = true;
            });
        var book = await bookBuilder.Build();

        await Verify(book);
    }
}