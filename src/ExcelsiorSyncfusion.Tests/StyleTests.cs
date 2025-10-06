// ReSharper disable UnusedParameter.Local

using Syncfusion.Drawing;
using Syncfusion.XlsIO;

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
                style.Font.Color = ExcelKnownColors.White;
                style.Color = Color.DarkBlue;
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
                style.Font.Color = ExcelKnownColors.White;
                style.Color = Color.DarkBlue;
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
                            style.Font.Color = ExcelKnownColors.Dark_green;
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
                        if (isActive)
                        {
                            style.Color = Color.LightGreen;
                        }
                        else
                        {
                            style.Color = Color.LightPink;
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