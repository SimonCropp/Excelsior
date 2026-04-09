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