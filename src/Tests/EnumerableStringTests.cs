[TestFixture]
public class EnumerableStringTests
{
    public record Person(string Name, string[] PhoneNumbers);

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithToString

        List<Person> data =
        [
            new("John Doe",
                PhoneNumbers:
                [
                    "+1 3057380950",
                    "+1 5056169368",
                    "+1 8634446859"
                ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }


    public record Target(string Value1, string[] Value2);

    [Test]
    public async Task WithNewlines()
    {
        #region ComplexTypeWithToString

        List<Target> data =
        [
            new("Value1",
                Value2:
                [
                    "a",
                    "a\n",
                    "a\r",
                    "a\r\n",
                    "a",
                    "a\na",
                    "a\ra",
                    "a\r\na",
                    "\ra",
                    "\na",
                    "\r\na",
                ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WithWhitespace()
    {
        #region ComplexTypeWithToString

        List<Target> data =
        [
            new("Value1",
                Value2:
                [
                    "a ",
                    "a  ",
                    " a ",
                    "  a  ",
                    "a\t",
                    "\ta",
                    "\ta\t",
                    "\na\n",
                ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WithWhitespaceNoTrim()
    {
        #region ComplexTypeWithToString

        List<Target> data =
        [
            new("Value1",
                Value2:
                [
                    "a ",
                    "a  ",
                    " a ",
                    "  a  ",
                    "a\t",
                    "\ta",
                    "\ta\t",
                    "\na\n",
                ]),
        ];

        var builder = new BookBuilder(trimWhitespace: false);
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WithNewlinesNoTrim()
    {
        #region ComplexTypeWithToString

        List<Target> data =
        [
            new("Value1",
                Value2:
                [
                    "a",
                    "a\n",
                    "a\r",
                    "a\r\n",
                    "a",
                    "a\na",
                    "a\ra",
                    "a\r\na",
                    "\ra",
                    "\na",
                    "\r\na",
                ]),
        ];

        var builder = new BookBuilder(trimWhitespace: false);
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}