[TestFixture]
public class Enums
{
    [Test]
    public async Task CustomRender()
    {
        var employees = SampleData.Employees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Status,
                _ => _.Render = (_, value) => $"Status: {value}");

        var book = await builder.Build();

        await Verify(book);
    }
}