// ReSharper disable ArrangeObjectCreationWhenTypeNotEvident
[TestFixture]
public class ComplexTypeWithSplitterUseHierachyForName
{
    public enum State
    {
        SA
    }

    public record Person(
        string Name,
        [Split(UseHierachyForName = true)]
        Address Address);

    public record Address(int Number, string Street, string City, State State, ushort PostCode);

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithSplitterUseHierachyForName

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