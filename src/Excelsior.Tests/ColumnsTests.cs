// ReSharper disable NotAccessedPositionalProperty.Local
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Local
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
        var columns = new Columns<NoOrderModel>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["First", "Second", "Third"]));
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
        var columns = new Columns<MixedOrderModel>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["Ordered", "NoOrder1", "NoOrder2"]));
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
        var columns = new Columns<AllOrderedModel>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["First", "Second", "Third"]));
    }

    record NoOrderRecord(string First, string Second, string Third);

    [Test]
    public void Record_NoOrder_UsesDeclarationOrder()
    {
        var columns = new Columns<NoOrderRecord>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["First", "Second", "Third"]));
    }

    record MixedOrderRecord(
        string NoOrder1,
        [Column(Order = 5)] string Ordered,
        string NoOrder2);

    [Test]
    public void Record_MixedOrder_UnorderedMaintainDeclarationPosition()
    {
        var columns = new Columns<MixedOrderRecord>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["Ordered", "NoOrder1", "NoOrder2"]));
    }

    record MixedConstructorAndProperties(string First, string Second)
    {
        public string Third { get; init; } = "";
        public string Fourth { get; init; } = "";
    }

    [Test]
    public void Record_MixedConstructorAndProperties_UsesDeclarationOrder()
    {
        var columns = new Columns<MixedConstructorAndProperties>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["First", "Second", "Third", "Fourth"]));
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
        var columns = new Columns<MixedConstructorAndPropertiesWithOrder>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["Ordered2", "Ordered1", "NoOrder1", "NoOrder2"]));
    }

    record AllOrderedRecord(
        [Column(Order = 3)] string Third,
        [Column(Order = 1)] string First,
        [Column(Order = 2)] string Second);

    [Test]
    public void Record_AllOrdered_SortsByOrder()
    {
        var columns = new Columns<AllOrderedRecord>();
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["First", "Second", "Third"]));
    }

    [Test]
    public void Fluent_ReorderColumns()
    {
        var columns = new Columns<NoOrderModel>();
        columns.Add<string>(_ => _.Third, _ => _.Order = 1);
        columns.Add<string>(_ => _.First, _ => _.Order = 2);
        columns.Add<string>(_ => _.Second, _ => _.Order = 3);
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["Third", "First", "Second"]));
    }

    [Test]
    public void Fluent_PartialOrder_UnorderedAfterOrdered()
    {
        var columns = new Columns<NoOrderModel>();
        columns.Add<string>(_ => _.Third, _ => _.Order = 1);
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["Third", "First", "Second"]));
    }

    class AttributeOrderModel
    {
        [Column(Order = 2)]
        public string A { get; set; } = "";

        public string B { get; set; } = "";

        [Column(Order = 1)]
        public string C { get; set; } = "";
    }

    [Test]
    public void Fluent_OverridesAttributeOrder()
    {
        var columns = new Columns<AttributeOrderModel>();
        columns.Add<string>(_ => _.B, _ => _.Order = 0);
        var ordered = columns.OrderedColumns();

        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["B", "C", "A"]));
    }

    [Test]
    public void Fluent_MixedWithAttributeAndPositional()
    {
        var columns = new Columns<MixedOrderModel>();
        // MixedOrderModel: NoOrder1 (no attr), Ordered (Order=5), NoOrder2 (no attr)
        // Fluent sets NoOrder2 to Order=3
        columns.Add<string>(_ => _.NoOrder2, _ => _.Order = 3);
        var ordered = columns.OrderedColumns();

        // Ordered(5), NoOrder2(3) are explicitly ordered; NoOrder1 is positional
        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["NoOrder2", "Ordered", "NoOrder1"]));
    }

    record FluentRecordModel(string A, string B, string C, string D);

    [Test]
    public void Fluent_Record_MixOrderedAndPositional()
    {
        var columns = new Columns<FluentRecordModel>();
        columns.Add<string>(_ => _.C, _ => _.Order = 1);
        columns.Add<string>(_ => _.A, _ => _.Order = 2);
        var ordered = columns.OrderedColumns();

        // C(1), A(2) explicitly ordered; B, D positional
        Assert.That(ordered.Select(_ => _.Name).ToList(), Is.EqualTo(["C", "A", "B", "D"]));
    }
}
