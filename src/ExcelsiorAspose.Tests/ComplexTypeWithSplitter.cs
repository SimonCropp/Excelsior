// ReSharper disable ArrangeObjectCreationWhenTypeNotEvident
[TestFixture]
public class ComplexTypeWithSplitter
{
    public enum State
    {
        SA
    }

    public record Person(string Name, [Split] Address Address);

    public record Address(int Number, string Street, string City, State State, ushort PostCode);
    //
    // #region ComplexTypeWithCustomRenderInit
    //
    // [ModuleInitializer]
    // public static void Init() =>
    //     ValueRenderer.For<Address>(_ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");
    //
    // #endregion

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
}