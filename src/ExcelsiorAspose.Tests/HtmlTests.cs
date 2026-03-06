[TestFixture]
public class HtmlTests
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
            _ => _.IsHtml = true);
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
            _ => _.IsHtml = true);
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
                _.IsHtml = true;
            });
        var book = await bookBuilder.Build();

        await Verify(book);
    }

    [Test]
    public async Task RichHtml()
    {
        var bookBuilder = new BookBuilder();

        List<Target> data =
        [
            new(
                """
                <h3>Status Report</h3>
                <p><b>Priority:</b> <span style="color: red">Critical</span></p>
                <ol>
                  <li>Review <code>PR #123</code></li>
                  <li>Fix <span style="color: green"><b>passing</b></span> tests</li>
                </ol>
                <p><i>Note:</i> See <a href="https://example.com">details</a></p>
                <table>
                  <tr><th>Region</th><th>Sales</th></tr>
                  <tr><td>North</td><td>$500K</td></tr>
                </table>
                """),
        ];
        var sheetBuilder = bookBuilder.AddSheet(data);

        sheetBuilder.Column(
            _ => _.Value1,
            _ => _.IsHtml = true);
        var book = await bookBuilder.Build();

        await Verify(book);
    }
}