[TestFixture]
public class ValueRendererForEnums
{
    #region ValueRendererForEnumsInit

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
                Manufacturer = Manufacturer.BuildYourDream,
                Color = Color.AntiqueWhite,
            }
        ];
        builder.AddSheet(data);

        #endregion

        using var book = await builder.Build();

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
        BuildYourDream
    }
}