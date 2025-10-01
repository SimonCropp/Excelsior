// ReSharper disable UnusedParameter.Local
[TestFixture]
public class Tests
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
    public async Task Render()
    {
        var employees = SampleData.Employees();

        #region AsposeCustomRender

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ => _.Render = (employee, name) => name.ToUpper())
            .Column(
                _ => _.IsActive,
                _ => _.Render = (employee, isActive) => isActive ? "Active" : "Inactive")
            .Column(
                _ => _.HireDate,
                _ => _.Format = "yyyy-MM-dd");

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task WorksheetName()
    {
        var employees = SampleData.Employees();

        #region AsposeWorksheetName

        var builder = new BookBuilder();
        builder.AddSheet(employees, "Employee Report");

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnWidths()
    {
        var employees = SampleData.Employees();

        #region AsposeColumnWidths

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ => _.Width = 25)
            .Column(_ => _.Email, _ => _.Width = 30)
            .Column(_ => _.HireDate, _ => _.Width = 15);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task Enums()
    {
        var employees = SampleData.Employees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Status,
                _ => _.Render = (_, value) => $"Status: {value}");

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task TypeSafeConfiguration()
    {
        var employees = SampleData.Employees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ =>
                {
                    _.Heading = "Full Name";
                    _.Width = 20;
                })
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Format = "$#,##0.00";
                    _.HeadingStyle = _ => _.Font.Color = Color.Green;
                });

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

    [Test]
    public async Task DisplayAttributes()
    {
        var builder = new BookBuilder();
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
        builder.AddSheet(products);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ToStream()
    {
        var data = SampleData.Employees();

        #region AsposeToStream

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var stream = new MemoryStream();
        await builder.ToStream(stream);

        #endregion

        await Verify(stream, extension: "xlsx");
    }


    [Test]
    public async Task RealWorldScenario()
    {
        var employees = GetRealWorldEmployeeData();
        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: Color.AliceBlue,
            headingStyle: style =>
            {
                style.Font.IsBold = true;
                style.Font.Color = Color.White;
                style.BackgroundColor = Color.DarkBlue;
                style.HorizontalAlignment = TextAlignmentType.Center;
            });
        builder.AddSheet(employees, "Employee Report 2024")
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.Format = "$#,##0.00";
                    config.CellStyle = (style, _, salary) =>
                    {
                        if (salary >= 100000)
                        {
                            style.BackgroundColor = Color.LightGreen;
                        }
                        else if (salary < 50000)
                        {
                            style.BackgroundColor = Color.LightPink;
                        }
                    };
                })
            .Column(
                _ => _.HireDate,
                config =>
                {
                    config.Format = "MMM dd, yyyy";
                    config.Width = 15;
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.Render = (_, value) => value ? "Active" : "Inactive";
                    config.CellStyle = (style, _, _) =>
                    {
                        style.HorizontalAlignment = TextAlignmentType.Center;
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
                                style.BackgroundColor = Color.PaleGreen;
                                break;
                            case EmployeeStatus.PartTime:
                                style.BackgroundColor = Color.LightYellow;
                                break;
                            case EmployeeStatus.Contract:
                                style.BackgroundColor = Color.LightCyan;
                                break;
                            case EmployeeStatus.Terminated:
                                style.BackgroundColor = Color.MistyRose;
                                break;
                        }
                    };
                });

        var book = await builder.Build();

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
            Salary = 125000,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1002,
            Name = "John Matrix",
            Email = "j.matrix@techcorp.com",
            HireDate = new(2020, 7, 22),
            Salary = 95000,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1003,
            Name = "Ellen Ripley",
            Email = "e.ripley@techcorp.com",
            HireDate = new(2019, 11, 8),
            Salary = 110000,
            IsActive = true,
            Status = EmployeeStatus.FullTime
        },
        new()
        {
            Id = 1004,
            Name = "Dutch Schaefer",
            Email = "d.schaefer@techcorp.com",
            HireDate = new(2021, 2, 14),
            Salary = 75000,
            IsActive = true,
            Status = EmployeeStatus.Contract
        },
        new()
        {
            Id = 1005,
            Name = "Kyle Reese",
            Email = "k.reese@techcorp.com",
            HireDate = new(2022, 6, 30),
            Salary = 45000,
            IsActive = false,
            Status = EmployeeStatus.PartTime
        },
        new()
        {
            Id = 1006,
            Name = "Roy Batty",
            Email = "r.batty@techcorp.com",
            HireDate = new(2017, 12, 1),
            Salary = 140000,
            IsActive = false,
            Status = EmployeeStatus.Terminated
        }
    ];

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