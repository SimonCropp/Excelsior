[TestFixture]
public class SpreadsheetAnchorTests
{
    [Test]
    public Task SimpleLink() =>
        Verify(HtmlToInlineString.Convert("<a href=\"https://example.com\">Example</a>"));

    [Test]
    public Task LinkWithSameText() =>
        Verify(HtmlToInlineString.Convert("<a href=\"https://example.com\">https://example.com</a>"));

    [Test]
    public Task LinkWithFormatting() =>
        Verify(HtmlToInlineString.Convert("<a href=\"https://example.com\"><b>Bold Link</b></a>"));

    [Test]
    public Task LinkInText() =>
        Verify(HtmlToInlineString.Convert("Visit <a href=\"https://example.com\">our site</a> for more info."));

    [Test]
    public Task LinkWithNoHref() =>
        Verify(HtmlToInlineString.Convert("<a>anchor text</a>"));

    [Test]
    public Task MultipleLinks() =>
        Verify(HtmlToInlineString.Convert(
            "<a href=\"https://one.com\">One</a> and <a href=\"https://two.com\">Two</a>"));
}
