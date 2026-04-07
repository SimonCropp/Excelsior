[TestFixture]
public class WorksheetName
{
    [Test]
    public async Task Fluent()
    {
        var employees = SampleData.Employees();

        #region WorksheetName

        var builder = new BookBuilder();
        builder.AddSheet(employees, "Employee Report");

        #endregion

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }
}