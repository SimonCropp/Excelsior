[TestFixture]
public class SpreadsheetBasicTests
{
    [Test]
    public Task PlainText() =>
        Verify(HtmlToInlineString.Convert("Hello world"));

    [Test]
    public Task Bold() =>
        Verify(HtmlToInlineString.Convert("<b>bold text</b>"));

    [Test]
    public Task Strong() =>
        Verify(HtmlToInlineString.Convert("<strong>strong text</strong>"));

    [Test]
    public Task Italic() =>
        Verify(HtmlToInlineString.Convert("<i>italic text</i>"));

    [Test]
    public Task Em() =>
        Verify(HtmlToInlineString.Convert("<em>emphasized</em>"));

    [Test]
    public Task Underline() =>
        Verify(HtmlToInlineString.Convert("<u>underlined</u>"));

    [Test]
    public Task Strikethrough() =>
        Verify(HtmlToInlineString.Convert("<s>struck</s>"));

    [Test]
    public Task StrikeTag() =>
        Verify(HtmlToInlineString.Convert("<strike>struck</strike>"));

    [Test]
    public Task Del() =>
        Verify(HtmlToInlineString.Convert("<del>deleted</del>"));

    [Test]
    public Task Superscript() =>
        Verify(HtmlToInlineString.Convert("x<sup>2</sup>"));

    [Test]
    public Task Subscript() =>
        Verify(HtmlToInlineString.Convert("H<sub>2</sub>O"));

    [Test]
    public Task LineBreak() =>
        Verify(HtmlToInlineString.Convert("line one<br>line two"));

    [Test]
    public Task SelfClosingBreak() =>
        Verify(HtmlToInlineString.Convert("line one<br/>line two"));

    [Test]
    public Task MixedFormatting() =>
        Verify(HtmlToInlineString.Convert("normal <b>bold</b> <i>italic</i> normal"));

    [Test]
    public Task EmptyHtml() =>
        Verify(HtmlToInlineString.Convert(""));

    [Test]
    public Task WhitespaceOnly() =>
        Verify(HtmlToInlineString.Convert("   "));

    [Test]
    public Task HtmlEntities() =>
        Verify(HtmlToInlineString.Convert("&amp; &lt; &gt; &quot; &apos;"));

    [Test]
    public Task NonBreakingSpace() =>
        Verify(HtmlToInlineString.Convert("hello&nbsp;world"));

    [Test]
    public Task InsTag() =>
        Verify(HtmlToInlineString.Convert("<ins>inserted</ins>"));
}
