// ReSharper disable ArrangeObjectCreationWhenTypeNotEvident
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

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task Null()
    {
        var builder = new BookBuilder();

        List<Person> data =
        [
            new("John Doe", null!)
        ];
        builder.AddSheet(data);

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }
}