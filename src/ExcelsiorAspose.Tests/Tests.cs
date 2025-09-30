[TestFixture]
public class Tests
{
    [Test]
    public async Task Simple()
    {
        #region AsposeUsage

        var builder = new BookBuilder();

        List<Employee> data =
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
            }
        ];
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }
    [Test]
    public async Task Whitespace()
    {
        #region AsposeWhitespace

        var builder = new BookBuilder();

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John Doe   ",
                Email = "    john@company.com    ",
            }
        ];
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }
    [Test]
    public async Task DisableWhitespaceTrim()
    {
        #region AsposeDisableWhitespaceTrim

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John Doe   ",
                Email = "    john@company.com    ",
            }
        ];

        var builder = new BookBuilder(trimWhitespace: false);
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }

    [Test]
    public async Task Nulls()
    {
        var builder = new BookBuilder();
        List<NullableTargets> data =
        [
            new()
            {
                Number = null,
                String = null,
                DateTime = null,
                Enum = null,
                Bool = null
            },
            new()
            {
                Number = 1,
                String = "value",
                DateTime = new DateTime(2020, 1, 1),
                Enum = AnEnum.Value,
                Bool = true
            },
        ];
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullsWithOverride()
    {
        var builder = new BookBuilder();
        List<NullableTargets> data =
        [
            new()
            {
                Number = null,
                String = null,
                DateTime = null,
                Enum = null,
                Bool = null
            },
            new()
            {
                Number = 1,
                String = "value",
                DateTime = new DateTime(2020, 1, 1),
                Enum = AnEnum.Value,
                Bool = true
            },
        ];
        var sheet = builder.AddSheet(data);
        sheet.Column(
            _ => _.Number,
            _ =>
            {
                _.CellStyle = (style, value) =>
                {
                    Debug.WriteLine(style);
                    Debug.WriteLine(value);
                };
                _.Render = _ => _.ToString();
            });
        sheet.Column(
            _ => _.DateTime,
            _ =>
            {
                _.CellStyle = (style, value) =>
                {
                    Debug.WriteLine(style);
                    Debug.WriteLine(value);
                };
                _.Render = _ =>
                {
                    if (_.HasValue)
                    {
                        return _.Value.ToString(CultureInfo.InvariantCulture);
                    }

                    return null;
                };
            });
        sheet.Column(
            _ => _.Enum,
            _ =>
            {
                _.CellStyle = (style, value) =>
                {
                    Debug.WriteLine(style);
                    Debug.WriteLine(value);
                };
                _.Render = _ => _.ToString();
            });
        sheet.Column(
            _ => _.String,
            _ =>
            {
                _.CellStyle = (style, value) =>
                {
                    Debug.WriteLine(style);
                    Debug.WriteLine(value);
                };
                _.Render = _ => _?.ToString();
            });
        sheet.Column(
            _ => _.Bool,
            _ =>
            {
                _.CellStyle = (style, value) =>
                {
                    Debug.WriteLine(style);
                    Debug.WriteLine(value);
                };
                _.Render = _ => _?.ToString().ToUpper();
            });
        var book = await builder.Build();

        await Verify(book);
    }

    public enum AnEnum
    {
        Value
    }

    public class NullableTargets
    {
        public required int? Number { get; init; }
        public required string? String { get; init; }
        public required DateTime? DateTime { get; init; }
        public required AnEnum? Enum { get; init; }
        public required bool? Bool { get; init; }
    }

    [Test]
    public async Task CustomHeaders()
    {
        var employees = GetSampleEmployees();

        #region AsposeCustomHeaders

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ => _.Header = "Employee Name");

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnOrdering()
    {
        var employees = GetSampleEmployees();

        #region AsposeColumnOrdering

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(_ => _.Email, _ => _.Order = 1)
            .Column(_ => _.Name, _ => _.Order = 2)
            .Column(_ => _.Salary, _ => _.Order = 3);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task HeaderStyle()
    {
        var data = GetSampleEmployees();

        #region AsposeHeaderStyle

        var builder = new BookBuilder(
            headerStyle: style =>
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
        var data = GetSampleEmployees();

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
        var employees = GetSampleEmployees();

        #region AsposeConditionalStyling

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.CellStyle = (style, value) =>
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
                    config.CellStyle = (style, isActive) =>
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
        var employees = GetSampleEmployees();

        #region AsposeCustomRender

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ => _.Render = value => value.ToUpper())
            .Column(
                _ => _.IsActive,
                _ => _.Render = active => active ? "Active" : "Inactive")
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
        var employees = GetSampleEmployees();

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
        var employees = GetSampleEmployees();

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

    public class ModelWithNulls
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Email { get; set; }
        public DateTime? HireDate { get; set; }
        public EmployeeStatus? Status { get; set; }
    }

    [Test]
    public async Task NullValues()
    {
        var builder = new BookBuilder();
        var employees = new List<ModelWithNulls>
        {
            new()
            {
                Id = 1,
                Name = "John Doe",
                Email = null,
                HireDate = new DateTime(2020, 1, 15),
                Status = EmployeeStatus.Contract,
            },
            new()
            {
                Id = 2,
                Name = null,
                Email = "jane@company.com",
                HireDate = null,
                Status = EmployeeStatus.Contract,
            },
            new()
            {
                Id = 3,
                Name = "Bob Johnson",
                Email = "bob@company.com",
                HireDate = new DateTime(2021, 7, 10),
                Status = EmployeeStatus.PartTime,
            }
        };
        builder.AddSheet(employees)
            .Column(_ => _.Name, _ => _.NullDisplay = "[No Name]")
            .Column(_ => _.HireDate, _ => _.NullDisplay = "[No HireDate]")
            .Column(_ => _.Email, _ => _.NullDisplay = "[No Email]")
            .Column(_ => _.Status, _ => _.NullDisplay = "[No Status]");

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullValuesShortcuts()
    {
        var bookBuilder = new BookBuilder();
        var employees = new List<ModelWithNulls>
        {
            new()
            {
                Id = 1,
                Name = "John Doe",
                Email = null,
                HireDate = new DateTime(2020, 1, 15),
                Status = EmployeeStatus.Contract,
            },
            new()
            {
                Id = 2,
                Name = null,
                Email = "jane@company.com",
                HireDate = null,
                Status = EmployeeStatus.Contract,
            },
            new()
            {
                Id = 3,
                Name = "Bob Johnson",
                Email = "bob@company.com",
                HireDate = new DateTime(2021, 7, 10),
                Status = EmployeeStatus.PartTime,
            }
        };
        var sheetBuilder = bookBuilder.AddSheet(employees);
        sheetBuilder.NullDisplayText(_ => _.Name, "[No Name]");
        sheetBuilder.NullDisplayText(_ => _.HireDate, "[No HireDate]");
        sheetBuilder.NullDisplayText(_ => _.Email, "[No Email]");
        sheetBuilder.NullDisplayText(_ => _.Status, "[No Status]");

        var book = await bookBuilder.Build();

        await Verify(book);
    }

    [Test]
    public async Task Enums()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Status,
                _ => _.Render = enumValue => $"Status: {enumValue}");

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task TypeSafeConfiguration()
    {
        var employees = GetSampleEmployees();
        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Name,
                _ =>
                {
                    _.Header = "Full Name";
                    _.Width = 20;
                })
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Format = "$#,##0.00";
                    _.HeaderStyle = _ => _.Font.Color = Color.Green;
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
        var data = GetSampleEmployees();

        #region AsposeToStream

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var stream = new MemoryStream();
        await builder.ToStream(stream);

        #endregion

        await Verify(stream, extension: "xlsx");
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

    [Test]
    public async Task RealWorldScenario()
    {
        var employees = GetRealWorldEmployeeData();
        var builder = new BookBuilder(
            useAlternatingRowColors: true,
            alternateRowColor: Color.AliceBlue,
            headerStyle: style =>
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
                    config.CellStyle = (style, salary) =>
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
                    config.Render = active => active ? "Active" : "Inactive";
                    config.CellStyle = (style, _) =>
                    {
                        style.HorizontalAlignment = TextAlignmentType.Center;
                    };
                })
            .Column(
                _ => _.Status,
                config =>
                {
                    config.CellStyle = (style, status) =>
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