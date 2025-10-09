// ReSharper disable UseSymbolAlias
[TestFixture]
public class DateFormats
{
    #region DateFormatsInit

    [ModuleInitializer]
    public static void UseHumanizerForEnums()
    {
        ValueRenderer.DefaultDateFormat = "yyyy/MM/dd" ;
        ValueRenderer.DefaultDateTimeFormat = "yyyy/MM/dd HH:mm:ss" ;
        ValueRenderer.DefaultDateTimeOffsetFormat = "yyyy/MM/dd HH:mm:ss z" ;
    }

    #endregion

    [Test]
    public async Task Test()
    {
        var builder = new BookBuilder();

        List<Model> data =
        [
            new()
            {
                DateProperty = new(2020, 10, 1),
                DateTimeProperty = new(2020, 10, 1, 6, 30, 0),
                DateTimeOffsetProperty = new(2020, 10, 1, 6, 30, 0, TimeSpan.FromHours(11)),
            }
        ];
        builder.AddSheet(data);

        using var book = await builder.Build();

        await Verify(book);
    }

    class Model
    {
        public DateOnly DateProperty { get; set; }
        public DateTime DateTimeProperty { get; set; }
        public DateTimeOffset DateTimeOffsetProperty { get; set; }
    }
}