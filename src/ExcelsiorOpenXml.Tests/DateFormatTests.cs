[TestFixture]
public class DateFormatTests
{
    [Test]
    public async Task DatesShouldBeFormatted()
    {
        var data = new[]
        {
            new { Id = 1, Date = new DateTime(2020, 1, 15) },
            new { Id = 2, Date = new DateTime(2021, 5, 20) }
        };

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and check cell values
        using var wb = new XLWorkbook(tempPath);
        var sheet = wb.Worksheets.First();

        // Check that date cells have number format applied
        var dateCell1 = sheet.Cell(2, 2); // First data row, second column (Date)
        var dateCell2 = sheet.Cell(3, 2); // Second data row, second column (Date)

        // Cleanup
        File.Delete(tempPath);

        // Assert that format is applied and dates are displayed correctly
        Assert.That(dateCell1.Style.NumberFormat.Format, Is.Not.Empty);
        Assert.That(dateCell2.Style.NumberFormat.Format, Is.Not.Empty);
        Assert.That(dateCell1.GetFormattedString(), Does.Contain("2020"));
        Assert.That(dateCell2.GetFormattedString(), Does.Contain("2021"));
    }
}
