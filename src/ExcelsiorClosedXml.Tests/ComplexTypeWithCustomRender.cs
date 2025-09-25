﻿[TestFixture]
public class ComplexTypeWithCustomRender
{
    public enum State
    {
        SA
    }

    public record Person(string Name, Address Address);

    public record Address(int Number, string Street, string City, State State, ushort PostCode);

    [Test]
    public async Task Test()
    {
        // ReSharper disable once ArrangeObjectCreationWhenTypeNotEvident
        #region ComplexTypeWithCustomRender

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

        BookBuilder.RenderFor<Address>(
            _ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}