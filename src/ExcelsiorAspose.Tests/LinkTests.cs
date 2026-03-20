[TestFixture]
public class LinkTests
{
    #region LinkModel

    public record Target(string Name, Link? Link);

    #endregion

    [Test]
    public async Task SingleLink()
    {
        #region LinkUsage

        List<Target> data =
        [
            new("Example", new Link("Example", "https://example.com")),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task LinkUrlOnly()
    {
        List<Target> data =
        [
            new("Example", new Link("https://example.com")),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullLink()
    {
        List<Target> data =
        [
            new("Example", null),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    public record LinkListTarget(string Name, IEnumerable<Link> Links);

    [Test]
    public async Task LinkList()
    {
        List<LinkListTarget> data =
        [
            new("Example",
            [
                new Link("Google", "https://google.com"),
                new Link("GitHub", "https://github.com"),
            ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task LinkListSingle()
    {
        List<LinkListTarget> data =
        [
            new("Example",
            [
                new Link("Google", "https://google.com"),
            ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task LinkListEmpty()
    {
        List<LinkListTarget> data =
        [
            new("Example", []),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }
}
