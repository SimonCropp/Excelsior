[TestFixture]
public class Enums
{
    [Test]
    public async Task CustomRender()
    {
        var employees = SampleData.Employees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Status,
                _ => _.Render = (_, value) => $"Status: {value.ToString()}");

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    [Test]
    public async Task NullCustomRender()
    {
        List<WithNull> employees =
        [
            new()
            {
                Value = null
            },
            new()
            {
                Value = AnEnum.Foo
            },
        ];
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Value,
                _ => _.Render = (_, value) => $"Value: {value}");

        using var stream = await builder.Build();

        await Verify(stream, "xlsx");
    }

    enum AnEnum
    {
        Foo
    };

    class WithNull
    {
        public AnEnum? Value { get; init; }
    }
}