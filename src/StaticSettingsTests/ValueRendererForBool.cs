[TestFixture]
public class ValueRendererForBool
{
    [SetUp]
    public void Setup() =>
        ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");

    [TearDown]
    public void Teardown() =>
        ValueRenderer.Reset();

    #region ValueRendererForBoolInit

    static void CustomBoolRender() =>
        ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");

    #endregion

    [Test]
    public async Task Test()
    {
        #region ValueRendererForBool

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
            }
        ];
        builder.AddSheet(data);

        #endregion

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    class Target
    {
        public string Name { get; set; } = null!;
        public bool IsActive { get; set; }
    }
}
