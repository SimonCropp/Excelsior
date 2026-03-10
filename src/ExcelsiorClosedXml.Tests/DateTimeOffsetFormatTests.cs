using System.Threading.Tasks;

public class DateTimeOffsetFormatTests
{
    [Test]
    public async Task CellStyle_UsesDateTimeOffsetFormat()
    {
        var builder = new BookBuilder();

        List<Model> data =
        [
            new()
            {
                DateTimeOffsetProperty = new(2020, 10, 1, 6, 30, 0, TimeSpan.FromHours(11)),
            }
        ];
        builder.AddSheet(data);

        using var book = await builder.Build();
        var sheet = book.Worksheets.First();
        var cell = sheet.Cell(2, 1); // row 2 = first data row, col 1 = first column

        await Assert.That(cell.Style.DateFormat.Format).IsEqualTo(ValueRenderer.DefaultDateTimeOffsetFormat);
    }

    [Test]
    public async Task CellStyle_WithCustomFormat_UsesCustomFormat()
    {
        var builder = new BookBuilder();

        List<Model> data =
        [
            new()
            {
                DateTimeOffsetProperty = new(2020, 10, 1, 6, 30, 0, TimeSpan.FromHours(11)),
            }
        ];
        var sheet = builder.AddSheet(data);
        sheet.Format(_ => _.DateTimeOffsetProperty, "yyyy/MM/dd HH:mm");

        using var book = await builder.Build();
        var worksheet = book.Worksheets.First();
        var cell = worksheet.Cell(2, 1);

        await Assert.That(cell.Style.DateFormat.Format).IsEqualTo("yyyy/MM/dd HH:mm");
    }

    class Model
    {
        public DateTimeOffset DateTimeOffsetProperty { get; set; }
    }
}