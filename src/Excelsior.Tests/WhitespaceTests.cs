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

        using var stream = await builder.Build();

        #endregion

        await Verify(stream, "xlsx");
    }

}