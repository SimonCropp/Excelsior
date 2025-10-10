// ReSharper disable ArrangeObjectCreationWhenTypeNotEvident
[TestFixture]
public class ComplexTypeWithSplitter
{
    public enum State
    {
        SA
    }

    public record Person(
        string Name,
        [Split]
        Address Address);

    public record Address(int Number, string Street, string City, State State, ushort PostCode);

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithSplitter

        var builder = new BookBuilder();

        List<Person> data =
        [
            new("John Doe",
                new Address(
                    Number: 900,
                    Street: "Victoria Square",
                    City: "Adelaide",
                    State: State.SA,
                    PostCode: 5000)),
        ];
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    public record Target(
        string Property,
        [Split]
        Target.NestedTarget Nested)
    {
        public record NestedTarget(string NestedProperty);
    }

    [Test]
    public async Task NestedColumn()
    {
        var builder = new BookBuilder();

        List<Target> data =
        [
            new("the Property",
                new Target.NestedTarget(
                    NestedProperty: "the NestedProperty")),
        ];
        var sheet = builder.AddSheet(data);
        sheet.Column(_ => _.Nested.NestedProperty, _ => _.Heading = "Custom");
        var book = await builder.Build();

        await Verify(book);
    }

    public record TargetMultipleOverlappingNested(
        string Property,
        [Split]
        TargetMultipleOverlappingNested.NestedTarget Nested1,
        [Split]
        TargetMultipleOverlappingNested.NestedTarget Nested2)
    {
        public record NestedTarget(string NestedProperty);
    }

    [Test]
    public async Task MultipleOverlappingNested()
    {
        var builder = new BookBuilder();

        List<TargetMultipleOverlappingNested> data =
        [
            new("the Property",
                new TargetMultipleOverlappingNested.NestedTarget(
                    NestedProperty: "the NestedProperty 1"),
                new TargetMultipleOverlappingNested.NestedTarget(
                    NestedProperty: "the NestedProperty 2")),
        ];
        builder.AddSheet(data);
        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task MultipleOverlappingNestedAndColumnConfig()
    {
        var builder = new BookBuilder();

        List<TargetMultipleOverlappingNested> data =
        [
            new("the Property",
                new TargetMultipleOverlappingNested.NestedTarget(
                    NestedProperty: "the NestedProperty 1"),
                new TargetMultipleOverlappingNested.NestedTarget(
                    NestedProperty: "the NestedProperty 2")),
        ];
        var sheet = builder.AddSheet(data);
        sheet.Column(_ => _.Nested1.NestedProperty, _ => _.Heading = "Custom1");
        sheet.Column(_ => _.Nested2.NestedProperty, _ => _.Heading = "Custom2");
        var book = await builder.Build();

        await Verify(book);
    }
}