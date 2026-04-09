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
}