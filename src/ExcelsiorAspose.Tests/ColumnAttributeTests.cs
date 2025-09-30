[TestFixture]
public class ColumnAttributeTests
{
    #region ColumnAttributeModels

    public class Employee
    {
        [Column(Header = "Employee ID", Order = 1)]
        public required int Id { get; init; }

        [Column(Header = "Full Name", Order = 2)]
        public required string Name { get; init; }

        [Column(Header = "Email Address")]
        public required string Email { get; init; }

        [Column(Header = "Hire Date", Order = 3)]
        public DateTime? HireDate { get; init; }
    }

    #endregion

    [Test]
    public async Task Test()
    {
        #region ColumnAttribute

        var builder = new BookBuilder();

        List<Employee> data =
        [
            new()
            {
                Id = 1,
                Name = "John Doe",
                Email = "john@company.com",
                HireDate = new(2020, 1, 15),
            }
        ];

        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}