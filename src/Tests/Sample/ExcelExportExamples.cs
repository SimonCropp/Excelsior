public static class ExcelExportExamples
{
    public static void BasicExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder();
        builder.AddSheet(employees);

        using var stream = new FileStream("employees.xlsx", FileMode.Create);
        builder.ToStream(stream);
    }

    public static void AdvancedExport()
    {
        var employees = GetSampleEmployees();

        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: XLColor.AliceBlue,
            headerStyle:style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            });
        builder.AddSheet(employees,"Employee Report")
            .ConfigureColumn(nameof(Employee.Salary), config =>
            {
                config.NumberFormat = "#,##0.00";
                config.HeaderStyle = style =>
                {
                    style.Fill.BackgroundColor = XLColor.Green;
                };
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is decimal and > 100000)
                    {
                        style.Font.FontColor = XLColor.DarkGreen;
                        style.Font.Bold = true;
                    }
                };
            })
            .ConfigureColumn(
                nameof(Employee.HireDate),
                config =>
            {
                config.DateTimeFormat = "yyyy-MM-dd";
                config.ColumnWidth = 15;
            })
            .ConfigureColumn(
                nameof(Employee.IsActive),
                config =>
            {
                config.BooleanDisplayFormat = active => active ? "✓ Yes" : "✗ No";
                config.DataCellStyle = style =>
                {
                    style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                };
            })
            .ConfigureColumn(
                nameof(Employee.Status),
                config =>
            {
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is EmployeeStatus status)
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
                    }
                };
            });

        using var stream = new FileStream("advanced_employees.xlsx", FileMode.Create);
        builder.ToStream(stream);
    }

    static List<Employee> GetSampleEmployees() =>
    [
        new()
        {
            Id = 1,
            Name = "John Doe",
            Email = "john@company.com",
            HireDate = new(2020, 1, 15),
            Salary = 75000m,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 2,
            Name = "Jane Smith",
            Email = "jane@company.com",
            HireDate = new(2019, 3, 22),
            Salary = 120000m,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 3,
            Name = "Bob Johnson",
            Email = "bob@company.com",
            HireDate = new(2021, 7, 10),
            Salary = 45000m,
            IsActive = false,
            Status = EmployeeStatus.PartTime
        },
        new()
        {
            Id = 4,
            Name = "Alice Brown",
            Email = "alice@company.com",
            HireDate = new(2018, 11, 5),
            Salary = 95000m,
            IsActive = true,
            Status = EmployeeStatus.Contract
        }
    ];
}