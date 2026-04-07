[TestFixture]
public class SpreadsheetHeadingTests
{
    [Test]
    public Task H1() =>
        Verify(HtmlToInlineString.Convert("<h1>Main Title</h1>"));

    [Test]
    public Task H2() =>
        Verify(HtmlToInlineString.Convert("<h2>Subtitle</h2>"));

    [Test]
    public Task H3() =>
        Verify(HtmlToInlineString.Convert("<h3>Section</h3>"));

    [Test]
    public Task H4() =>
        Verify(HtmlToInlineString.Convert("<h4>Subsection</h4>"));

    [Test]
    public Task H5() =>
        Verify(HtmlToInlineString.Convert("<h5>Minor</h5>"));

    [Test]
    public Task H6() =>
        Verify(HtmlToInlineString.Convert("<h6>Smallest</h6>"));

    [Test]
    public Task HeadingWithInlineFormatting() =>
        Verify(HtmlToInlineString.Convert("<h1>Title with <i>italic</i> word</h1>"));

    [Test]
    public Task HeadingFollowedByParagraph() =>
        Verify(HtmlToInlineString.Convert("<h2>Heading</h2><p>Body text</p>"));
}
