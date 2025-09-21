[TestFixture]
public class ComplexTypeWithToString
{
    public enum State
    {
        SA
    }

    #region ComplexTypeModels

    public record Person(string Name, Address Address);

    public record Address(int Number, string Street, string City, State State, ushort PostCode);

    #endregion

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithToString

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

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}