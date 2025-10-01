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