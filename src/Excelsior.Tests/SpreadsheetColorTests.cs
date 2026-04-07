[TestFixture]
public class SpreadsheetColorTests
{
    [Test]
    public Task FontColorAttribute() =>
        Verify(HtmlToInlineString.Convert("<font color=\"#FF0000\">red text</font>"));

    [Test]
    public Task FontColorShortHex() =>
        Verify(HtmlToInlineString.Convert("<font color=\"#F00\">red text</font>"));

    [Test]
    public Task NamedColor() =>
        Verify(HtmlToInlineString.Convert("<span style=\"color: blue\">blue text</span>"));

    [Test]
    public Task RgbColor() =>
        Verify(HtmlToInlineString.Convert("<span style=\"color: rgb(0, 128, 0)\">green text</span>"));

    [Test]
    public Task MultipleColors() =>
        Verify(HtmlToInlineString.Convert(
            """
            <span style="color: red">red</span>
            <span style="color: blue">blue</span>
            <span style="color: green">green</span>
            """));

    [Test]
    public Task ColorWithFormatting() =>
        Verify(HtmlToInlineString.Convert(
            "<b style=\"color: #FF0000\">bold red</b>"));

    [Test]
    public Task NestedColors() =>
        Verify(HtmlToInlineString.Convert(
            "<span style=\"color: red\">outer <span style=\"color: blue\">inner</span> outer</span>"));
}
