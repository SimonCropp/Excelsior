// ReSharper disable UnusedParameter.Local
[TestFixture]
public class FormulaTests
{
    [Test]
    public async Task Fluent()
    {
        var employees = SampleData.Employees();

        #region FormulaFluent

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Formula = (employee, context) =>
                        $"={context.Ref(_ => _.Id)} * 10000";
                    _.Format = "#,##0";
                });

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task SimpleOverload()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Formula(
                _ => _.Salary,
                context => $"={context.Ref(_ => _.Id)} * 1000");

        var book = await builder.Build();

        await Verify(book);
    }
}
