[TestFixture]
public class ComplexTypeWithInheritedCustomRender
{
    public enum State
    {
        SA
    }

    public record Person(string Name, IAddress Address);

    public interface IAddress
    {
        int Number { get; }
        string Street { get; }
        string City { get; }
        State State { get; }
        ushort PostCode { get; }
    }

    public record Address(int Number, string Street, string City, State State, ushort PostCode) :
        IAddress;

    [Test]
    public async Task Test()
    {
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

        ValueRenderer.For<IAddress>(
            _ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }
}