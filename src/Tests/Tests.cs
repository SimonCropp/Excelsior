using System.Globalization;

[TestFixture]
public class Tests
{
    [Test]
    public Task ClosedXmlFormatting()
    {
        List<string> items = [
            "Item1",
            "Item2",
            "Item with spaces",
            "Item\rwith\nnew\r\nlines",
            "Last"];
        using var book = new XLWorkbook();

        var sheet = book.Worksheets.Add();

        sheet.Cell("A1").Value = "ID";
        var cell = sheet.Cell("B1");
        cell.Style.Alignment.WrapText = true;
        var richText = cell.CreateRichText();
        foreach (var item in items)
        {
            richText.AddText("• ").SetBold();
            var indented = item
                .Replace("\r\n","\n")
                .Replace("\r","\n")
                .Replace("\n","\n   ");
            richText.AddText(indented);
            richText.AddNewLine();
        }

        sheet.Columns().AdjustToContents();

        return Verify(book);
    }

    [Test]
    public async Task Simple()
    {
        #region Usage

        List<Employee> data =
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
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }
    [Test]
    public async Task Whitespace()
    {
        #region Whitespace

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John\nDoe   ",
                Email = "    john@company.com    ",
            }
        ];

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        #endregion

        await Verify(book);
    }
    [Test]
    public async Task DisableWhitespaceTrim()
    {
        #region DisableWhitespaceTrim

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "    John\nDoe   ",
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

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task NullsWithOverride()
    {
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

        var builder = new BookBuilder();
        var sheet = builder.AddSheet(data);
        sheet.Column(
            _ => _.Number,
            _ =>
            {
                _.ConditionalStyling = (style, value) =>
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
                _.ConditionalStyling = (style, value) =>
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
                _.ConditionalStyling = (style, value) =>
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
                _.ConditionalStyling = (style, value) =>
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
                _.ConditionalStyling = (style, value) =>
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
        var data = GetSampleEmployees();

        #region CustomHeaders

        var builder = new BookBuilder();
        builder.AddSheet(data)
            .Column(
                _ => _.Name,
                _ => _.HeaderText = "Employee Name");

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnOrdering()
    {
        var data = GetSampleEmployees();

        #region ColumnOrdering

        var builder = new BookBuilder();
        builder.AddSheet(data)
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

        #region HeaderStyle

        var builder = new BookBuilder(
            headerStyle: style =>
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
        var data = GetSampleEmployees();

        #region GlobalStyle

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
    public async Task ConditionalStyling()
    {
        var employees = GetSampleEmployees();

        #region ConditionalStyling

        var builder = new BookBuilder();
        builder.AddSheet(employees)
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.ConditionalStyling = (style, value) =>
                    {
                        if (value > 100000)
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
                    config.ConditionalStyling = (style, isActive) =>
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
    public async Task Render()
    {
        var data = GetSampleEmployees();

        #region CustomRender

        var builder = new BookBuilder();
        builder.AddSheet(data)
            .Column(
                _ => _.Email,
                _ => _.Render = value => $"📧 {value}")
            .Column(
                _ => _.IsActive,
                _ => _.Render = active => active ? "✓ Active" : "✗ Inactive")
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

        #region WorksheetName

        var builder = new BookBuilder();
        builder.AddSheet(employees, "Employee Report");

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ColumnWidths()
    {
        var data = GetSampleEmployees();

        #region ColumnWidths

        var builder = new BookBuilder();
        builder.AddSheet(data)
            .Column(_ => _.Name, _ => _.ColumnWidth = 25)
            .Column(_ => _.Email, _ => _.ColumnWidth = 30)
            .Column(_ => _.HireDate, _ => _.ColumnWidth = 15);

        #endregion

        var book = await builder.Build();

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
            .Column(_ => _.Name, _ => _.NullDisplayText = "[No Name]")
            .Column(_ => _.Email, _ => _.NullDisplayText = "[No Email]");

        var book = await builder.Build();

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
                    _.HeaderText = "Full Name";
                    _.ColumnWidth = 20;
                })
            .Column(
                _ => _.Salary,
                _ =>
                {
                    _.Format = "$#,##0.00";
                    _.HeaderStyle = _ => _.Font.FontColor = XLColor.Green;
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

        var book = await builder.Build();

        await Verify(book);
    }

    [Test]
    public async Task ToStream()
    {
        var data = GetSampleEmployees();

        #region ToStream

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
        builder.AddSheet(employees, "Employee Report 2024")
            .Column(
                _ => _.Salary,
                config =>
                {
                    config.Format = "$#,##0.00";
                    config.ConditionalStyling = (style, salary) =>
                    {
                        if (salary >= 100000)
                        {
                            style.Fill.BackgroundColor = XLColor.LightGreen;
                        }
                        else if (salary < 50000)
                        {
                            style.Fill.BackgroundColor = XLColor.LightPink;
                        }
                    };
                })
            .Column(
                _ => _.HireDate,
                config =>
                {
                    config.Format = "MMM dd, yyyy";
                    config.ColumnWidth = 15;
                })
            .Column(
                _ => _.IsActive,
                config =>
                {
                    config.Render = active => active ? "✓ Active" : "✗ Inactive";
                    config.DataCellStyle = style =>
                    {
                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
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