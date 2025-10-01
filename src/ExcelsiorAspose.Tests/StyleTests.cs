// ReSharper disable UnusedParameter.Local
[TestFixture]
public class StyleTests
{
    [Test]
    public async Task HeadingStyle()
    {
        var data = SampleData.Employees();

        #region AsposeHeadingStyle

        var builder = new BookBuilder(
            headingStyle: style =>
            {
                style.Font.IsBold = true;
                style.Font.Color = Color.White;
                style.BackgroundColor = Color.DarkBlue;
            });
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task GlobalStyle()
    {
        var data = SampleData.Employees();

        #region AsposeGlobalStyle

        var builder = new BookBuilder(
            globalStyle: style =>
            {
                style.Font.IsBold = true;
                style.Font.Color = Color.White;
                style.BackgroundColor = Color.DarkBlue;
            });
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ConditionalStyling()
    {
        var employees = SampleData.Employees();

        #region AsposeConditionalStyling

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.CellStyle = (style, employee, value) =>
                    {
                        if (value > 100000)
                        {
                            style.Font.Color = Color.DarkGreen;
                            style.Font.IsBold = true;
                        }
                    };
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.CellStyle = (style, employee, isActive) =>
                    {
                        if (isActive)
                        {
                            style.BackgroundColor = Color.LightGreen;
                        }
                        else
                        {
                            style.BackgroundColor = Color.LightPink;
                        }
                    };
                });

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task EmptyList()
    {
        var builder = new BookBuilder();
        builder.AddSheet(new List<Employee>());

        var book = await builder.Build();

        await Verify(book);
    }
}