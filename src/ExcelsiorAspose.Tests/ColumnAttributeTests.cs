[TestFixture]
public class ColumnAttributeTests
{
    #region ColumnAttributeModel

    public class Employee
    {
        [Column(Heading = "Employee ID", Order = 1, Format = "0000")]
        public required int Id { get; init; }

        [Column(Heading = "Full Name", Order = 2, Width = 20)]
        public required string Name { get; init; }

        [Column(Heading = "Email Address", Width = 30)]
        public required string Email { get; init; }

        [Column(Heading = "Hire Date", Order = 3, NullDisplay = "unknown")]
        public Date? HireDate { get; init; }
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
            },
            new()
            {
                Id = 2,
                Name = "Jane Smith",
                Email = "jane@company.com",
                HireDate = null,
            }
        ];

        builder.AddSheet(data);

        #endregion

        var book = await builder.Build();

        await Verify(book);
    }
}