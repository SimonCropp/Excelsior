[TestFixture]
public class ColumnsTests
{
    class NoOrderModel
    {
        public string First { get; set; } = "";
        public string Second { get; set; } = "";
        public string Third { get; set; } = "";
    }

    [Test]
    public void NoOrder_UsesDeclarationOrder()
    {
        var columns = new Columns<NoOrderModel, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "First", "Second", "Third" }));
    }

    class MixedOrderModel
    {
        public string NoOrder1 { get; set; } = "";

        [Column(Order = 5)]
        public string Ordered { get; set; } = "";

        public string NoOrder2 { get; set; } = "";
    }

    [Test]
    public void MixedOrder_UnorderedMaintainDeclarationPosition()
    {
        var columns = new Columns<MixedOrderModel, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "NoOrder1", "NoOrder2", "Ordered" }));
    }

    class AllOrderedModel
    {
        [Column(Order = 3)]
        public string Third { get; set; } = "";

        [Column(Order = 1)]
        public string First { get; set; } = "";

        [Column(Order = 2)]
        public string Second { get; set; } = "";
    }

    [Test]
    public void AllOrdered_SortsByOrder()
    {
        var columns = new Columns<AllOrderedModel, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "First", "Second", "Third" }));
    }
}
