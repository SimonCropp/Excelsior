[TestFixture]
public class Headings
{
    [Test]
    public async Task Fluent()
    {
        var employees = SampleData.Employees();

        #region CustomHeadings

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ => _.Heading = "Employee Name");

        #endregion

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }
}