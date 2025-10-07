[TestFixture]
public class UsageTests
{
    #region ValueRendererForEnums

    [ModuleInitializer]
    public static void UseHumanizerForEnums() =>
        ValueRenderer.ForEnums(_ => _.Humanize());

    #endregion

    [Test]
    public async Task Test()
    {
        #region ValueRendererForEnums

        var builder = new BookBuilder();

        List<Car> data =
        [
            new()
            {
                Manufacturer = Manufacturer.BuildYouDream,
                Color = Color.AntiqueWhite,
            }
        ];
        builder.AddSheet(data);

        using var book = await builder.Build();

        #endregion

        await Verify(book);
    }

    enum Color
    {
        AntiqueWhite
    }

    class Car
    {
        public Manufacturer Manufacturer { get; set; }
        public Color Color { get; set; }
    }

    enum Manufacturer
    {
        BuildYouDream
    }
}