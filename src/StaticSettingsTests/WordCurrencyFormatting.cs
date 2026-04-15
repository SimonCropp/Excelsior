using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

[TestFixture]
public class WordCurrencyFormatting
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
    public void DefaultCultureProducesLocalCurrencySymbol()
    {
        ValueRenderer.Culture = CultureInfo.GetCultureInfo("en-US");

        var table = new WordTableBuilder<Item>(sample).Build();
        var priceText = ExtractPriceCellText(table);

        // en-US should produce "$1,234"; we assert on $ to avoid coupling to thousand-sep details.
        IsTrue(priceText.StartsWith('$'), $"expected en-US currency symbol, got '{priceText}'");
        AreEqual("$1,234", priceText);
    }

    [Test]
    public void CultureOverrideChangesCurrencySymbol()
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

    static string ExtractPriceCellText(Table table)
    {
        var dataRow = table.Elements<TableRow>().Skip(1).First();
        var priceCell = dataRow.Elements<TableCell>().Skip(1).First();
        return priceCell.GetFirstChild<Paragraph>()!.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text;
    }
}
