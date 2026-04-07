[TestFixture]
public class SpreadsheetMiscElementTests
{
    [Test]
    public Task AbbrTag() =>
        Verify(HtmlToInlineString.Convert(
            "The <abbr title=\"World Health Organization\">WHO</abbr> recommends it."));

    [Test]
    public Task AcronymTag() =>
        Verify(HtmlToInlineString.Convert(
            "Use <acronym title=\"HyperText Markup Language\">HTML</acronym> for web pages."));

    [Test]
    public Task TimeTag() =>
        Verify(HtmlToInlineString.Convert(
            "The meeting is at <time datetime=\"14:00\">2 PM</time>."));

    [Test]
    public Task QTag() =>
        Verify(HtmlToInlineString.Convert(
            "She said <q>hello world</q> to everyone."));

    [Test]
    public Task NestedQ() =>
        Verify(HtmlToInlineString.Convert(
            "<q>outer <q>inner</q> outer</q>"));

    [Test]
    public Task FigcaptionTag() =>
        Verify(HtmlToInlineString.Convert(
            "<figure><img alt=\"Chart\"><figcaption>Figure 1: Sales data</figcaption></figure>"));

    [Test]
    public Task SvgTag() =>
        Verify(HtmlToInlineString.Convert(
            "before <svg width=\"100\" height=\"100\"><circle cx=\"50\" cy=\"50\" r=\"40\"/></svg> after"));

    [Test]
    public Task ArticleTag() =>
        Verify(HtmlToInlineString.Convert(
            "<article>Article content here</article>"));

    [Test]
    public Task AsideTag() =>
        Verify(HtmlToInlineString.Convert(
            "<aside>Sidebar content</aside>"));

    [Test]
    public Task SectionTag() =>
        Verify(HtmlToInlineString.Convert(
            "<section>Section content</section>"));

    [Test]
    public Task DtWithBold() =>
        Verify(HtmlToInlineString.Convert(
            "<dl><dt>Term</dt><dd>Definition of the term</dd></dl>"));

    [Test]
    public Task BlockquoteWithQ() =>
        Verify(HtmlToInlineString.Convert(
            "<blockquote><q>To be or not to be</q></blockquote>"));
}
