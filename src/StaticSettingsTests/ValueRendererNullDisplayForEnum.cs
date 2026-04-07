[TestFixture]
public class ValueRendererNullDisplayForEnum
{
    [SetUp]
    public void Setup() =>
        ValueRenderer.NullDisplayFor<Enum>("Unknown");

    [TearDown]
    public void Teardown() =>
        ValueRenderer.Reset();

    #region ValueRendererNullDisplayForEnumInit

    static void CustomNullEnumDisplay() =>
        ValueRenderer.NullDisplayFor<Enum>("Unknown");

    #endregion

    [Test]
    public async Task Test()
    {
        #region ValueRendererNullDisplayForEnum

        var builder = new BookBuilder();

        List<Target> data =
        [
            new()
            {
                Name = "Alice",
                Color = Color.AntiqueWhite,
            },
            new()
            {
                Name = "Bob",
            }
        ];
        builder.AddSheet(data);

        #endregion

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    enum Color
    {
        AntiqueWhite
    }

    class Target
    {
        public string Name { get; set; } = null!;
        public Color? Color { get; set; }
    }
}
