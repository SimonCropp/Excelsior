[TestFixture]
public class EnumerableStringTests
{
    public record Person(string Name, string[] PhoneNumbers);

    [Test]
    public async Task Test()
    {
        #region ComplexTypeWithToString

        List<Person> data =
        [
            new("John Doe",
                PhoneNumbers:
                [
                    "+1 3057380950",
                    "+1 5056169368",
                    "+1 8634446859"
                ]),
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}