[TestFixture]
public class TypeInferenceTests
{
    public class InferenceModel
    {
        public required string Name { get; init; }
        public string? Notes { get; init; }
        public int Age { get; init; }
        public int? Score { get; init; }
        public bool IsActive { get; init; }
        public DateTime HireDate { get; init; }
    }

    [Test]
    public async Task TemplateInfersDefaults()
    {
        #region TemplateInferenceDefaults

        var builder = new BookBuilder();
        builder.AddTemplateSheet("Employees", templateRowCount: 10)
            .Column<string>("Name")
            .Column<int>("Age")
            .Column<bool>("IsActive")
            .Column<DateTime>("HireDate");

        using var book = await builder.Build();

        #endregion

        await Verify(book);
    }

    [Test]
    public async Task TemplateInferenceCanBeDisabled()
    {
        var builder = new BookBuilder();
        builder.AddTemplateSheet("Employees", inferValidationFromTypes: false)
            .Column<int>("Age")
            .Column<bool>("IsActive");

        using var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task TemplatePerColumnOptOut()
    {
        var builder = new BookBuilder();
        builder.AddTemplateSheet("Employees", templateRowCount: 5)
            .Column<int>(
                "Age",
                _ =>
                {
                    _.Required = false;
                })
            .Column<bool>(
                "IsActive",
                _ =>
                {
                    _.DisableAllowedValues = true;
                });

        using var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task DataBoundInfersWhenEnabled()
    {
        #region DataBoundInferenceEnabled

        InferenceModel[] data =
        [
            new()
            {
                Name = "Alice",
                Age = 30,
                IsActive = true,
                HireDate = new(2020, 1, 1)
            }
        ];

        var builder = new BookBuilder();
        builder.AddSheet(
            data,
            templateRowCount: 5,
            inferValidationFromTypes: true);

        using var book = await builder.Build();

        #endregion

        await Verify(book);
    }

    [Test]
    public async Task DataBoundDefaultsToNoInference()
    {
        InferenceModel[] data =
        [
            new()
            {
                Name = "Alice",
                Age = 30,
                IsActive = true,
                HireDate = new(2020, 1, 1)
            }
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        using var book = await builder.Build();

        await Verify(book);
    }
}
