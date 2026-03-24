using Humanizer;

[TestFixture]
public class ValueRendererForEnumsHumanizer
{
    #region ValueRendererForEnumsHumanizerInit

    [ModuleInitializer]
    public static void CustomEnumRender() =>
        ValueRenderer.ForEnums(_ => _.Humanize());

    #endregion

    [Test]
    public async Task Test()
    {
        #region ValueRendererForEnumsHumanizer

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
        public Color? NullableColor { get; set; }
    }

    enum Manufacturer
    {
        BuildYourDream
    }
}
