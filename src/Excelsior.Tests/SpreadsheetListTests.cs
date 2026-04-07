[TestFixture]
public class SpreadsheetListTests
{
    [Test]
    public Task UnorderedList() =>
        Verify(HtmlToInlineString.Convert("<ul><li>item one</li><li>item two</li><li>item three</li></ul>"));

    [Test]
    public Task OrderedList() =>
        Verify(HtmlToInlineString.Convert("<ol><li>first</li><li>second</li><li>third</li></ol>"));

    [Test]
    public Task SingleListItem() =>
        Verify(HtmlToInlineString.Convert("<ul><li>only item</li></ul>"));

    [Test]
    public Task FormattedListItems() =>
        Verify(HtmlToInlineString.Convert("<ul><li><b>bold item</b></li><li><i>italic item</i></li></ul>"));

    [Test]
    public Task NestedLists() =>
        Verify(HtmlToInlineString.Convert("<ul><li>outer</li><li><ul><li>inner</li></ul></li></ul>"));

    [Test]
    public Task DeeplyNestedLists() =>
        Verify(HtmlToInlineString.Convert("<ul><li>level 0</li><li><ul><li>level 1</li><li><ul><li>level 2</li></ul></li></ul></li></ul>"));

    [Test]
    public Task NestedOrderedList() =>
        Verify(HtmlToInlineString.Convert("<ol><li>first</li><li><ol><li>nested</li></ol></li></ol>"));
}
