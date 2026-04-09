[TestFixture]
public class ColumnWidths
{
    [Test]
    public async Task Fluent()
    {
        var employees = SampleData.Employees();

        #region ColumnWidths

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ => _.Width = 25)
            .Column(_ => _.Email, _ => _.Width = 30)
            .Column(_ => _.HireDate, _ => _.Width = 15);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task MinWidthFluent()
    {
        var employees = SampleData.Employees();

        #region ColumnMinWidth

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ => _.MinWidth = 40);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task MaxWidthFluent()
    {
        var employees = SampleData.Employees();

        #region ColumnMaxWidth

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ => _.MaxWidth = 5);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public void MinWidthEqualsMaxWidthThrows()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ =>
            {
                _.MinWidth = 25;
                _.MaxWidth = 25;
            });

        var exception = Assert.ThrowsAsync<Exception>(async () => await builder.Build());
        Assert.That(exception!.Message, Does.Contain("Use Width instead"));
    }

    [Test]
    public void MinWidthGreaterThanMaxWidthThrows()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ =>
            {
                _.MinWidth = 30;
                _.MaxWidth = 10;
            });

        var exception = Assert.ThrowsAsync<Exception>(async () => await builder.Build());
        Assert.That(exception!.Message, Does.Contain("MinWidth (30) is greater than MaxWidth (10)"));
    }

    #region ColumnMinMaxWidthModel

    public class EmployeeWithMinMaxWidth
    {
        [Column(MinWidth = 40)]
        public required string Name { get; init; }

        [Column(MaxWidth = 20)]
        public required string Email { get; init; }
    }

    #endregion
}