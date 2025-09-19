[TestFixture]
public class ComplexTypeWithToString
{
    public enum State
    {
        SouthAustralia
    }

    #region ComplexTypeModels

    public record Person(string Name, Address Address);

    public record Address(int Number, string Street, State State, string City, ushort PostCode);

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
                    State: State.SouthAustralia,
                    City: "Adelaide",
                    PostCode: 5000)),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = builder.Build();

        #endregion

        await Verify(book);
    }
}