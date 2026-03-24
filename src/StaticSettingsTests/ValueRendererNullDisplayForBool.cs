[TestFixture]
public class ValueRendererNullDisplayForBool
{
    [SetUp]
    public void Setup()
    {
        ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");
        ValueRenderer.NullDisplayFor<bool>("Unknown");
    }

    [TearDown]
    public void Teardown() =>
        ValueRenderer.Reset();

    #region ValueRendererNullDisplayForBoolInit

    static void CustomBoolRender()
    {
        ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");
        ValueRenderer.NullDisplayFor<bool>("Unknown");
    }

    #endregion

    [Test]
    public async Task Test()
    {
        #region ValueRendererNullDisplayForBool

        var builder = new BookBuilder();

        List<Target> data =
        [
            new()
            {
                Name = "Alice",
                IsActive = true,
            },
            new()
            {
                Name = "Bob",
                IsActive = false,
            },
            new()
            {
                Name = "Charlie",
                IsActive = null,
            }
        ];
        builder.AddSheet(data);

        #endregion

        using var book = await builder.Build();

        await Verify(book);
    }

    class Target
    {
        public string Name { get; set; } = null!;
        public bool? IsActive { get; set; }
    }
}
