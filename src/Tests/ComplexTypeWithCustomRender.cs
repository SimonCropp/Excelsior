[TestFixture]
public class ComplexTypeWithCustomRender
{
    public enum State
    {
        SouthAustralia
    }

    public record Person(string Name, Address Address);

    public record Address(int Number, string Street, State State, string City, ushort PostCode);

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithCustomRender

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

        BookBuilder.RenderFor<Address>(
            _ => $"{_.Number}, {_.Street}, {_.State}, {_.City}, {_.PostCode}");

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = builder.Build();

        #endregion

        await Verify(book);
    }
}