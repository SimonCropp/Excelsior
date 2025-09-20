[TestFixture]
public class ComplexTypeWithInheritedCustomRender
{
    public enum State
    {
        SouthAustralia
    }

    public record Person(string Name, IAddress Address);

    public interface IAddress
    {
        int Number { get; }
        string Street { get; }
        State State { get; }
        string City { get; }
        ushort PostCode { get; }
    }

    public record Address(int Number, string Street, State State, string City, ushort PostCode) : IAddress;

    [Test]
    public async Task Test()
    {
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

        BookBuilder.RenderFor<IAddress>(
            _ => $"{_.Number}, {_.Street}, {_.State}, {_.City}, {_.PostCode}");

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }
}