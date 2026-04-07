[TestFixture]
public class SpreadsheetBlockTests
{
    [Test]
    public Task Paragraphs() =>
        Verify(HtmlToInlineString.Convert("<p>first paragraph</p><p>second paragraph</p>"));

    [Test]
    public Task Divs() =>
        Verify(HtmlToInlineString.Convert("<div>first div</div><div>second div</div>"));

    [Test]
    public Task Headings() =>
        Verify(HtmlToInlineString.Convert("<h1>heading one</h1><h2>heading two</h2><h3>heading three</h3>"));

    [Test]
    public Task MixedBlocksAndInline() =>
        Verify(HtmlToInlineString.Convert("<p>paragraph with <b>bold</b></p><div>div with <i>italic</i></div>"));

    [Test]
    public Task Blockquote() =>
        Verify(HtmlToInlineString.Convert("<blockquote>quoted text</blockquote>"));

    [Test]
    public Task PreformattedText() =>
        Verify(HtmlToInlineString.Convert("<pre>  line one\n  line two</pre>"));

    [Test]
    public Task HorizontalRule() =>
        Verify(HtmlToInlineString.Convert("above<hr>below"));

    [Test]
    public Task DefinitionList() =>
        Verify(HtmlToInlineString.Convert("<dl><dt>Term</dt><dd>Definition</dd></dl>"));
}
