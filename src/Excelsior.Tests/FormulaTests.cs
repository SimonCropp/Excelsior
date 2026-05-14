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
                    _.Width = 15;
                });

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public void FormulaWithoutWidthThrows()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Formula(_ => _.Salary, context => $"={context.Ref(_ => _.Id)} * 1000");

        var exception = Assert.ThrowsAsync<Exception>(() => builder.Build());
        Assert.That(exception!.Message, Does.Contain("formula columns must set Width explicitly"));
    }

    [Test]
    public void FormulaWithMinWidthThrows()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Formula = (employee, context) => $"={context.Ref(_ => _.Id)} * 1000";
                    _.MinWidth = 10;
                });

        var exception = Assert.ThrowsAsync<Exception>(() => builder.Build());
        Assert.That(exception!.Message, Does.Contain("formula columns cannot use MinWidth/MaxWidth"));
    }

    [Test]
    public void FormulaWithMaxWidthThrows()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Formula = (employee, context) => $"={context.Ref(_ => _.Id)} * 1000";
                    _.MaxWidth = 30;
                });

        var exception = Assert.ThrowsAsync<Exception>(() => builder.Build());
        Assert.That(exception!.Message, Does.Contain("formula columns cannot use MinWidth/MaxWidth"));
    }

    [Test]
    public async Task SimpleOverload()
    {
        var employees = SampleData.Employees();

        var builder = new BookBuilder();
        var sheet = builder.AddSheet(employees);
        sheet.Formula(_ => _.Salary, context => $"={context.Ref(_ => _.Id)} * 1000");
        sheet.Width(_ => _.Salary, 15);

        var book = await builder.Build();

        await Verify(book);
    }
}
