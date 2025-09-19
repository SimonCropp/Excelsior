[TestFixture]
public class Tests
{
    [Test]
    public async Task Simple()
    {
        #region Usage

        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees);

        var book = builder.Build();

        #endregion
        await Verify(book);
    }

    [Test]
    public async Task CustomHeaders()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(
                nameof(Employee.Name),
                _ => _.HeaderText = "Employee Name")
            .ConfigureColumn(
                nameof(Employee.Email),
                _ => _.HeaderText = "Email Address");

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnOrdering()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Email), _ => _.Order = 1)
            .ConfigureColumn(nameof(Employee.Name), _ => _.Order = 2)
            .ConfigureColumn(nameof(Employee.Salary), _ => _.Order = 3);

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task Styling()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: XLColor.LightGray,
            headerStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
            });
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Salary), config =>
            {
                config.NumberFormat = "#,##0.00";
                config.HeaderStyle = style =>
                {
                    style.Fill.BackgroundColor = XLColor.Green;
                };
            });

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task GlobalStyling()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: XLColor.LightGray,
            globalStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
            });
        builder.AddSheet(employees);

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ConditionalStyling()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Salary), config =>
            {
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is decimal and > 100000)
                    {
                        style.Font.FontColor = XLColor.DarkGreen;
                        style.Font.Bold = true;
                    }
                };
            })
            .ConfigureColumn(nameof(Employee.IsActive), config =>
            {
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is bool isActive)
                    {
                        style.Fill.BackgroundColor = isActive ? XLColor.LightGreen : XLColor.LightPink;
                    }
                };
            });

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task CustomFormatters()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Email), _ => _.CustomFormatter = value => $"📧 {value}")
            .ConfigureColumn(nameof(Employee.IsActive), _ => _.BooleanDisplayFormat = active => active ? "✓ Active" : "✗ Inactive")
            .ConfigureColumn(nameof(Employee.HireDate), _ => _.DateTimeFormat = "yyyy-MM-dd");

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WorksheetName()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees, "Employee Report");

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnWidths()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Name), _ => _.ColumnWidth = 25)
            .ConfigureColumn(nameof(Employee.Email), _ => _.ColumnWidth = 30)
            .ConfigureColumn(nameof(Employee.HireDate), _ => _.ColumnWidth = 15);

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullValues()
    {
        var employees = new List<EmployeeWithNulls>
        {
            new()
            {
                Id = 1,
                Name = "John Doe",
                Email = null,
                HireDate = new DateTime(2020, 1, 15)
            },
            new()
            {
                Id = 2,
                Name = null,
                Email = "jane@company.com",
                HireDate = null
            },
            new()
            {
                Id = 3,
                Name = "Bob Johnson",
                Email = "bob@company.com",
                HireDate = new DateTime(2021, 7, 10)
            }
        };

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(EmployeeWithNulls.Name), _ => _.NullDisplayText = "[No Name]")
            .ConfigureColumn(nameof(EmployeeWithNulls.Email), _ => _.NullDisplayText = "[No Email]");

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task Enums()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(nameof(Employee.Status), _ => _.EnumDisplayFormat = enumValue => $"Status: {enumValue}");

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task TypeSafeConfiguration()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .ConfigureColumn(_ => _.Name, _ =>
            {
                _.HeaderText = "Full Name";
                _.ColumnWidth = 20;
            })
            .ConfigureColumn(_ => _.Salary, _ =>
            {
                _.NumberFormat = "$#,##0.00";
                _.HeaderStyle = _ => _.Font.FontColor = XLColor.Green;
            });

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task EmptyList()
    {
        var employees = new List<Employee>();
        var builder = new BookBuilder();
        builder.AddSheet(employees);

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task DisplayAttributes()
    {
        var products = new List<Product>
        {
            new()
            {
                Id = 1,
                Name = "Laptop",
                Price = 999.99m,
                Category = ProductCategory.Electronics,
                IsAvailable = true
            },
            new()
            {
                Id = 2,
                Name = "Book",
                Price = 29.99m,
                Category = ProductCategory.Books,
                IsAvailable = false
            }
        };

        var builder = new BookBuilder();
        builder.AddSheet(products);

        var book = builder.Build();

        await Verify(book);
    }

    [Test]
    public Task ToStream()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees);

        var stream = new MemoryStream();
        builder.ToStream(stream);
        return Verify(stream, extension: "xlsx");
    }

    [Test]
    public void InvalidPropertyName()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        var converter = builder.AddSheet(employees);

        Assert.DoesNotThrow(() =>
            converter.ConfigureColumn("NonExistentProperty", _ =>
            {
                _.HeaderText = "Test";
            })
        );
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

    [Test]
    public async Task RealWorldScenario()
    {
        var employees = GetRealWorldEmployeeData();
        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: XLColor.AliceBlue,
            headerStyle: style =>
            {
                style.Font.Bold = true;
                style.Font.FontColor = XLColor.White;
                style.Fill.BackgroundColor = XLColor.DarkBlue;
                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            });
        builder.AddSheet(employees,"Employee Report 2024")
            .ConfigureColumn(nameof(Employee.Salary), config =>
            {
                config.NumberFormat = "$#,##0.00";
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is decimal salary)
                    {
                        if (salary >= 100000)
                        {
                            style.Fill.BackgroundColor = XLColor.LightGreen;
                        }
                        else if (salary < 50000)
                        {
                            style.Fill.BackgroundColor = XLColor.LightPink;
                        }
                    }
                };
            })
            .ConfigureColumn(nameof(Employee.HireDate), config =>
            {
                config.DateTimeFormat = "MMM dd, yyyy";
                config.ColumnWidth = 15;
            })
            .ConfigureColumn(nameof(Employee.IsActive), config =>
            {
                config.BooleanDisplayFormat = active => active ? "✓ Active" : "✗ Inactive";
                config.DataCellStyle = style =>
                {
                    style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                };
            })
            .ConfigureColumn(nameof(Employee.Status), config =>
            {
                config.ConditionalStyling = (style, value) =>
                {
                    if (value is EmployeeStatus status)
                    {
                        switch (status)
                        {
                            case EmployeeStatus.FullTime:
                                style.Fill.BackgroundColor = XLColor.PaleGreen;
                                break;
                            case EmployeeStatus.PartTime:
                                style.Fill.BackgroundColor = XLColor.LightYellow;
                                break;
                            case EmployeeStatus.Contract:
                                style.Fill.BackgroundColor = XLColor.LightCyan;
                                break;
                            case EmployeeStatus.Terminated:
                                style.Fill.BackgroundColor = XLColor.MistyRose;
                                style.Font.Strikethrough = true;
                                break;
                        }
                    }
                };
            });

        var book = builder.Build();

        await Verify(book);
    }

    static List<Employee> GetRealWorldEmployeeData() =>
    [
        new()
        {
            Id = 1001,
            Name = "Sarah Connor",
            Email = "s.connor@techcorp.com",
            HireDate = new(2018, 3, 15),
            Salary = 125000m,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1002,
            Name = "John Matrix",
            Email = "j.matrix@techcorp.com",
            HireDate = new(2020, 7, 22),
            Salary = 95000m,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1003,
            Name = "Ellen Ripley",
            Email = "e.ripley@techcorp.com",
            HireDate = new(2019, 11, 8),
            Salary = 110000m,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1004,
            Name = "Dutch Schaefer",
            Email = "d.schaefer@techcorp.com",
            HireDate = new(2021, 2, 14),
            Salary = 75000m,
            IsActive = true,
            Status = EmployeeStatus.Contract
        },
        new()
        {
            Id = 1005,
            Name = "Kyle Reese",
            Email = "k.reese@techcorp.com",
            HireDate = new(2022, 6, 30),
            Salary = 45000m,
            IsActive = false,
            Status = EmployeeStatus.PartTime
        },
        new()
        {
            Id = 1006,
            Name = "Roy Batty",
            Email = "r.batty@techcorp.com",
            HireDate = new(2017, 12, 1),
            Salary = 140000m,
            IsActive = false,
            Status = EmployeeStatus.Terminated
        }
    ];

    public class EmployeeWithNulls
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Email { get; set; }
        public DateTime? HireDate { get; set; }
    }

    public class Product
    {
        [Display(Name = "Product ID", Order = 1)]
        public required int Id { get; set; }

        [Display(Name = "Product Name", Order = 2)]
        public required string Name { get; set; }

        [Display(Name = "Unit Price", Order = 3)]
        public required decimal Price { get; set; }

        [DisplayName("Product Category")]
        public required ProductCategory Category { get; set; }

        public required bool IsAvailable { get; set; }
    }

    public enum ProductCategory
    {
        [Display(Name = "Electronic Items")]
        Electronics,

        [Display(Name = "Books & Literature")]
        Books,

        [Display(Name = "Clothing & Apparel")]
        Clothing,

        [Display(Name = "Home & Garden")]
        HomeGarden
    }
}