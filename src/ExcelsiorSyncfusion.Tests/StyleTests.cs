// ReSharper disable UnusedParameter.Local
[TestFixture]
public class StyleTests
{
    [Test]
    public async Task HeadingStyle()
    {
        var data = SampleData.Employees();

        #region SyncfusionHeadingStyle

        var builder = new BookBuilder(
            headingStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
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

        #region SyncfusionGlobalStyle

        var builder = new BookBuilder(
            globalStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
            });
        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task CellStyle()
    {
        var employees = SampleData.Employees();

        #region SyncfusionCellStyle

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.CellStyle = (style, employee, salary) =>
                    {
                        if (salary > 100000)
                        {
                            style.Font.FontColor = XLColor.DarkGreen;
                            style.Font.Bold = true;
                        }
                    };
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.CellStyle = (style, employee, isActive) =>
                    {
                        var fill = style.Fill;
                        if (isActive)
                        {
                            fill.BackgroundColor = XLColor.LightGreen;
                        }
                        else
                        {
                            fill.BackgroundColor = XLColor.LightPink;
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