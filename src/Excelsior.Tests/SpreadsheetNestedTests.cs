[TestFixture]
public class SpreadsheetNestedTests
{
    [Test]
    public Task BoldItalic() =>
        Verify(HtmlToInlineString.Convert("<b><i>bold italic</i></b>"));

    [Test]
    public Task BoldUnderlineItalic() =>
        Verify(HtmlToInlineString.Convert("<b><u><i>all three</i></u></b>"));

    [Test]
    public Task NestedSameTag() =>
        Verify(HtmlToInlineString.Convert("<b>outer <b>inner</b> outer</b>"));

    [Test]
    public Task PartialOverlap() =>
        Verify(HtmlToInlineString.Convert("<b>bold <i>bold-italic</i> bold</b>"));

    [Test]
    public Task DeeplyNested() =>
        Verify(HtmlToInlineString.Convert("<b><i><u><s>all formats</s></u></i></b>"));

    [Test]
    public Task MixedContent() =>
        Verify(HtmlToInlineString.Convert("start <b>bold <i>both</i></b> <u>under</u> end"));
}
