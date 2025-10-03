public static class ExcelExportExamples
{
    public static async Task BasicExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder();
        builder.AddSheet(employees);

        await using var stream = new FileStream("employees.xlsx", FileMode.Create);
        await builder.ToStream(stream);
    }

    public static async Task AdvancedExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: XLColor.AliceBlue,
            headingStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            });
        builder.AddSheet(employees, "Employee Report")
            .Column(_ => _.Salary,
                config =>
                {
                    config.Format = "#,##0.00";
                    config.HeadingStyle = style =>
                        style.Fill.BackgroundColor = XLColor.Green;
                    config.CellStyle = (style, _, value) =>
                    {
                        if (value > 100000)
                        {
                            style.Font.FontColor = XLColor.DarkGreen;
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
                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
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
                                style.Fill.BackgroundColor = XLColor.LightGreen;
                                break;
                            case EmployeeStatus.PartTime:
                                style.Fill.BackgroundColor = XLColor.LightYellow;
                                break;
                            case EmployeeStatus.Contract:
                                style.Fill.BackgroundColor = XLColor.LightBlue;
                                break;
                            case EmployeeStatus.Terminated:
                                style.Fill.BackgroundColor = XLColor.LightPink;
                                break;
                        }
                    };
                });

        await using var stream = new FileStream("advanced_employees.xlsx", FileMode.Create);
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