[TestFixture]
public class WhitespaceTests
{
    [Test]
    public async Task Whitespace()
    {
        #region Whitespace

        var builder = new BookBuilder();

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John Doe   ",
                Email = "    john@company.com    ",
            }
        ];
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }

    [Test]
    public async Task Disable()
    {
        #region DisableWhitespaceTrim

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John Doe   ",
                Email = "    john@company.com    ",
            }
        ];

        var builder = new BookBuilder(trimWhitespace: false);
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }

}