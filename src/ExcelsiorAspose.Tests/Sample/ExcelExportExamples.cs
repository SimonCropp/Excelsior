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
            headerStyle: style =>
            {
                style.Font.IsBold = true;
                style.Font.Color =  Color.White;
                style.BackgroundColor = Color.DarkBlue;
                style.HorizontalAlignment = TextAlignmentType.Center;
            });
        builder.AddSheet(employees, "Employee Report")
            .Column(_ => _.Salary,
                config =>
                {
                    config.Format = "#,##0.00";
                    config.HeaderStyle = style =>
                    {
                        style.BackgroundColor = Color.Green;
                    };
                    config.ConditionalStyling = (style, value) =>
                    {
                        if (value > 100000)
                        {
                            style.Font.Color = Color.DarkGreen;
                            style.Font.IsBold = true;
                        }
                    };
                })
            .Column(
                _ => _.HireDate,
                config =>
                {
                    config.Format = "yyyy-MM-dd";
                    config.ColumnWidth = 15;
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.Render = active => active ? "✓ Yes" : "✗ No";
                    config.DataCellStyle = style =>
                    {
                        style.HorizontalAlignment = TextAlignmentType.Center;
                    };
                })
            .Column(
                _ => _.Status,
                config =>
                {
                    config.ConditionalStyling = (style, status) =>
                    {
                        switch (status)
                        {
                            case EmployeeStatus.FullTime:
                                style.BackgroundColor = Color.LightGreen;
                                break;
                            case EmployeeStatus.PartTime:
                                style.BackgroundColor = Color.LightYellow;
                                break;
                            case EmployeeStatus.Contract:
                                style.BackgroundColor = Color.LightBlue;
                                break;
                            case EmployeeStatus.Terminated:
                                style.BackgroundColor = Color.LightPink;
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