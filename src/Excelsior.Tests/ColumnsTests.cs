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

    record NoOrderRecord(string First, string Second, string Third);

    [Test]
    public void Record_NoOrder_UsesDeclarationOrder()
    {
        var columns = new Columns<NoOrderRecord, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "First", "Second", "Third" }));
    }

    record MixedOrderRecord(
        string NoOrder1,
        [Column(Order = 5)] string Ordered,
        string NoOrder2);

    [Test]
    public void Record_MixedOrder_UnorderedMaintainDeclarationPosition()
    {
        var columns = new Columns<MixedOrderRecord, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "NoOrder1", "NoOrder2", "Ordered" }));
    }

    record MixedConstructorAndProperties(string First, string Second)
    {
        public string Third { get; init; } = "";
        public string Fourth { get; init; } = "";
    }

    [Test]
    public void Record_MixedConstructorAndProperties_UsesDeclarationOrder()
    {
        var columns = new Columns<MixedConstructorAndProperties, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "First", "Second", "Third", "Fourth" }));
    }

    record MixedConstructorAndPropertiesWithOrder(
        string NoOrder1,
        [Column(Order = 10)] string Ordered1)
    {
        public string NoOrder2 { get; init; } = "";

        [Column(Order = 5)]
        public string Ordered2 { get; init; } = "";
    }

    [Test]
    public void Record_MixedConstructorAndPropertiesWithOrder()
    {
        var columns = new Columns<MixedConstructorAndPropertiesWithOrder, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "NoOrder1", "NoOrder2", "Ordered2", "Ordered1" }));
    }

    record AllOrderedRecord(
        [Column(Order = 3)] string Third,
        [Column(Order = 1)] string First,
        [Column(Order = 2)] string Second);

    [Test]
    public void Record_AllOrdered_SortsByOrder()
    {
        var columns = new Columns<AllOrderedRecord, object>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(new[] { "First", "Second", "Third" }));
    }
}
