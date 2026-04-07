[TestFixture]
public class SpreadsheetStyleAttributeTests
{
    [Test]
    public Task FontWeightBold() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"font-weight: bold\">bold</span>"));

    [Test]
    public Task FontWeight700() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"font-weight: 700\">bold</span>"));

    [Test]
    public Task FontStyleItalic() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"font-style: italic\">italic</span>"));

    [Test]
    public Task TextDecorationUnderline() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"text-decoration: underline\">underlined</span>"));

    [Test]
    public Task TextDecorationLineThrough() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"text-decoration: line-through\">struck</span>"));

    [Test]
    public Task MultipleStyleProperties() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"font-weight: bold; font-style: italic; color: #FF0000\">styled</span>"));

    [Test]
    public Task VerticalAlignSuper() =>
        Verify(HtmlToInlineString.Convert(
            "E = mc<span style=\"vertical-align: super\">2</span>"));

    [Test]
    public Task VerticalAlignSub() =>
        Verify(HtmlToInlineString.Convert(
            "H<span style=\"vertical-align: sub\">2</span>O"));
}
