using Syncfusion.Drawing;
using Syncfusion.XlsIO;

public static class ExcelExportExamples
{
    public static async Task BasicExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder();
        builder.AddSheet(employees);

        using var stream = new FileStream("employees.xlsx", FileMode.Create);
        await builder.ToStream(stream);
    }

    public static async Task AdvancedExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: Color.AliceBlue,
            headingStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.Color = ExcelKnownColors.White;
                style.Color =Color.DarkBlue;
                style.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            });
        builder.AddSheet(employees, "Employee Report")
            .Column(_ => _.Salary,
                config =>
                {
                    config.Format = "#,##0.00";
                    config.HeadingStyle = style =>
                    {
                        style.Color = Color.Green;
                    };
                    config.CellStyle = (style, _, value) =>
                    {
                        if (value > 100000)
                        {
                            style.Font.Color = ExcelKnownColors.Dark_green;
                            style.Font.Bold = true;
                        }
                    };
                })
            .Column(
                _ => _.HireDate,
                config =>
                {
                    config.Format = "yyyy-MM-dd";
                    config.Width = 15;
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.Render = (_, active) => active ? "Yes" : "No";
                    config.CellStyle = (style, _, _) =>
                    {
                        style.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    };
                })
            .Column(
                _ => _.Status,
                config =>
                {
                    config.CellStyle = (style, _, status) =>
                    {
                        switch (status)
                        {
                            case EmployeeStatus.FullTime:
                                style.Color = Color.LightGreen;
                                break;
                            case EmployeeStatus.PartTime:
                                style.Color = Color.LightYellow;
                                break;
                            case EmployeeStatus.Contract:
                                style.Color = Color.LightBlue;
                                break;
                            case EmployeeStatus.Terminated:
                                style.Color = Color.LightPink;
                                break;
                        }
                    };
                });

        using var stream = new FileStream("advanced_employees.xlsx", FileMode.Create);
        await builder.ToStream(stream);
    }

    static List<Employee> GetSampleEmployees() =>
    [
        new()
        {
            Id = 1,
            Name = "John Doe",
            Email = "john@company.com",
            HireDate = new(2020, 1, 15),
            Salary = 75000,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 2,
            Name = "Jane Smith",
            Email = "jane@company.com",
            HireDate = new(2019, 3, 22),
            Salary = 120000,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 3,
            Name = "Bob Johnson",
            Email = "bob@company.com",
            HireDate = new(2021, 7, 10),
            Salary = 45000,
            IsActive = false,
            Status = EmployeeStatus.PartTime
        },
        new()
        {
            Id = 4,
            Name = "Alice Brown",
            Email = "alice@company.com",
            HireDate = new(2018, 11, 5),
            Salary = 95000,
            IsActive = true,
            Status = EmployeeStatus.Contract
        }
    ];
}