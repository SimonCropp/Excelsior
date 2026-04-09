[TestFixture]
public class RowHeights
{
    public class Note
    {
        public required string Title { get; init; }
        public required string Body { get; init; }
    }

    static List<Note> Notes() =>
    [
        new()
        {
            Title = "Short",
            Body = "A brief note."
        },
        new()
        {
            Title = "Long",
            Body = "This is a much longer note that contains a lot of text and is expected to wrap across many visual lines when rendered in a narrow column, which means the row would otherwise grow very tall."
        },
        new()
        {
            Title = "Multiline",
            Body = "Line one.\nLine two.\nLine three.\nLine four.\nLine five.\nLine six.\nLine seven."
        }
    ];

    [Test]
    public async Task MaxRowHeightFluent()
    {
        var notes = Notes();

        #region MaxRowHeight

        var builder = new BookBuilder();
        builder.AddSheet(notes, maxRowHeight: 60);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task BookMaxRowHeight()
    {
        var notes = Notes();

        #region BookMaxRowHeight

        var builder = new BookBuilder(maxRowHeight: 60);
        builder.AddSheet(notes);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task MaxRowHeightFluentZeroRows()
    {
        var builder = new BookBuilder();
        builder.AddSheet(new List<Note>(), maxRowHeight: 60);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task MaxRowHeightDoesNotShrinkSmallRows()
    {
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees(), maxRowHeight: 60);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public void MaxRowHeightBelowExcelDefaultThrows()
    {
        var builder = new BookBuilder();
        builder.AddSheet(Notes(), maxRowHeight: 14);

        var exception = Assert.ThrowsAsync<Exception>(async () => await builder.Build());
        Assert.That(exception!.Message, Does.Contain("MaxRowHeight (14) must be between 15 (the Excel default row height) and 409"));
    }

    [Test]
    public void MaxRowHeightExceedsExcelMaxThrows()
    {
        var builder = new BookBuilder();
        builder.AddSheet(Notes(), maxRowHeight: 410);

        var exception = Assert.ThrowsAsync<Exception>(async () => await builder.Build());
        Assert.That(exception!.Message, Does.Contain("MaxRowHeight (410) must be between 15 (the Excel default row height) and 409"));
    }
}
