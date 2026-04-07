[TestFixture]
public class SpreadsheetEdgeCaseTests
{
    [Test]
    public Task UnclosedTags() =>
        Verify(HtmlToInlineString.Convert("<b>bold <i>italic"));

    [Test]
    public Task ExtraClosingTags() =>
        Verify(HtmlToInlineString.Convert("text</b></i>more"));

    [Test]
    public Task ConsecutiveBreaks() =>
        Verify(HtmlToInlineString.Convert("one<br><br><br>two"));

    [Test]
    public Task WhitespaceCollapsing() =>
        Verify(HtmlToInlineString.Convert("  lots   of    spaces  "));

    [Test]
    public Task TabsAndNewlines() =>
        Verify(HtmlToInlineString.Convert("text\twith\ttabs\nand\nnewlines"));

    [Test]
    public Task SpecialCharacters() =>
        Verify(HtmlToInlineString.Convert("price: $100 & tax < 10% > 5%"));

    [Test]
    public Task UnknownTags() =>
        Verify(HtmlToInlineString.Convert("<custom>text</custom>"));

    [Test]
    public Task ImageAlt() =>
        Verify(HtmlToInlineString.Convert("before <img alt=\"image description\"> after"));

    [Test]
    public Task SpanWithNoStyle() =>
        Verify(HtmlToInlineString.Convert("<span>plain span</span>"));

    [Test]
    public Task MultipleSpaces() =>
        Verify(HtmlToInlineString.Convert("one     two"));

    [Test]
    public Task EmptyTags() =>
        Verify(HtmlToInlineString.Convert("<b></b><i></i>text"));

    [Test]
    public Task MalformedHtml() =>
        Verify(HtmlToInlineString.Convert("<b>bold <i>overlap</b> still italic</i>"));

    [Test]
    public Task NumericEntity() =>
        Verify(HtmlToInlineString.Convert("&#169; copyright"));

    [Test]
    public Task CiteTag() =>
        Verify(HtmlToInlineString.Convert("<cite>citation</cite>"));

    [Test]
    public Task DfnTag() =>
        Verify(HtmlToInlineString.Convert("<dfn>definition</dfn>"));

    [Test]
    public Task VarTag() =>
        Verify(HtmlToInlineString.Convert("<var>variable</var>"));

    [Test]
    public Task SampTag() =>
        Verify(HtmlToInlineString.Convert("<samp>sample output</samp>"));
}
