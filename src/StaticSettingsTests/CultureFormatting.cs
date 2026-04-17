using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using XlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using XlRow = DocumentFormat.OpenXml.Spreadsheet.Row;

[TestFixture]
public class CultureFormatting
{
    [TearDown]
    public void Teardown() =>
        ValueRenderer.Reset();

    public class Item
    {
        public required string Name { get; init; }

        [Column(Format = "C0")]
        public required decimal Price { get; init; }
    }

    static readonly Item[] sample = [new() { Name = "Widget", Price = 1234m }];

    [Test]
    public void WordCurrencyFormattingDefaultsToLocalCulture()
    {
        ValueRenderer.Culture = CultureInfo.GetCultureInfo("en-US");

        var table = new WordTableBuilder<Item>(sample).Build();
        var priceText = ExtractPriceCellText(table);

        // en-US should produce "$1,234"; we assert on $ to avoid coupling to thousand-sep details.
        IsTrue(priceText.StartsWith('$'), $"expected en-US currency symbol, got '{priceText}'");
        AreEqual("$1,234", priceText);
    }

    [Test]
    public void WordCurrencyFormattingHonorsCultureOverride()
    {
        ValueRenderer.Culture = CultureInfo.GetCultureInfo("en-GB");

        var table = new WordTableBuilder<Item>(sample).Build();
        var priceText = ExtractPriceCellText(table);

        AreEqual("£1,234", priceText);
    }

    [Test]
    public void ResetRestoresCurrentCulture()
    {
        ValueRenderer.Culture = CultureInfo.GetCultureInfo("en-GB");
        ValueRenderer.Reset();

        AreEqual(CultureInfo.CurrentCulture, ValueRenderer.Culture);
    }

    public class TimestampedItem
    {
        public required string Name { get; init; }

        [Column(Format = "MMMM dd, yyyy")]
        public required DateTimeOffset RecordedAt { get; init; }
    }

    [Test]
    public async Task ExcelDateTimeOffsetFormattingHonorsCulture()
    {
        // DateTimeOffset has no native Excel cell type, so it gets pre-formatted as an inline
        // string and the .NET-side culture controls month names. Switch to fr-FR and assert the
        // month appears in French.
        ValueRenderer.Culture = CultureInfo.GetCultureInfo("fr-FR");

        var items = new[]
        {
            new TimestampedItem
            {
                Name = "Sample",
                RecordedAt = new DateTimeOffset(2026, 1, 15, 0, 0, 0, TimeSpan.Zero)
            }
        };

        var builder = new BookBuilder();
        builder.AddSheet(items);
        using var book = await builder.Build();

        // Pull the inline-string text out of the data row's RecordedAt cell (column B, row 2).
        var sheetPart = book.WorkbookPart!.WorksheetParts.Single();
        var dataRow = sheetPart.Worksheet!.Descendants<XlRow>().Single(r => r.RowIndex?.Value == 2);
        var cell = dataRow.Elements<XlCell>().Single(c => c.CellReference?.Value == "B2");
        var rendered = cell.InlineString!.Text!.Text;

        AreEqual("janvier 15, 2026", rendered);
    }

    static string ExtractPriceCellText(Table table)
    {
        var dataRow = table.Elements<TableRow>().Skip(1).First();
        var priceCell = dataRow.Elements<TableCell>().Skip(1).First();
        return priceCell.GetFirstChild<Paragraph>()!.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text;
    }
}
