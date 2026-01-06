[TestFixture]
public class RowHeightTests
{
    public record Target(string LongText);

    [Test]
    public async Task RowHeightShouldNotExceedMaximum()
    {
        var bookBuilder = new BookBuilder();

        // Create a very long text that would cause row height to exceed 409 when wrapped
        var longText = string.Join("\n", Enumerable.Range(1, 100).Select(i => $"Line {i}: This is a very long line of text that will wrap"));

        List<Target> data =
        [
            new(longText),
        ];

        var sheetBuilder = bookBuilder.AddSheet(data);
        sheetBuilder.Column(_ => _.LongText, _ => _.Width = 30);

        var book = await bookBuilder.Build();

        // Verify that no row height exceeds the Excel maximum of 409
        var sheet = book.Worksheets.First();
        var lastRow = sheet.LastRowUsed()!.RowNumber();

        for (var i = 1; i <= lastRow; i++)
        {
            var height = sheet.Row(i).Height;
            Assert.That(height, Is.LessThanOrEqualTo(409),
                $"Row {i} height {height} exceeds maximum allowed height of 409");
        }
    }
}
