[TestFixture]
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
        var sheet = book.Worksheets[0];
        var cell = sheet.Cells[1, 0];
        var style = cell.GetStyle();

        Assert.That(style.Custom, Is.EqualTo(ValueRenderer.DefaultDateTimeOffsetFormat));
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
        var worksheet = book.Worksheets[0];
        var cell = worksheet.Cells[1, 0];
        var style = cell.GetStyle();

        Assert.That(style.Custom, Is.EqualTo("yyyy/MM/dd HH:mm"));
    }

    class Model
    {
        public DateTimeOffset DateTimeOffsetProperty { get; set; }
    }
}
