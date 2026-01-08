[TestFixture]
public class UsageTests
{
    [Test]
    public async Task Test()
    {
        var employees = SampleData.Employees();

        #region Usage

        var builder = new BookBuilder();
        builder.AddSheet(employees);

        #endregion

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and verify
        using var wb = new XLWorkbook(tempPath);
        await Verify(wb);

        // Cleanup
        File.Delete(tempPath);
    }
}
