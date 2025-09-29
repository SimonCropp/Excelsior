[TestFixture]
public class ColumnAttributeTests
{
    #region ColumnAttributeModels

    public class Employee
    {
        [Column(HeaderText = "Employee ID", Order = 1)]
        public required int Id { get; init; }

        [Column(HeaderText = "Full Name", Order = 2)]
        public required string Name { get; init; }

        [Column(HeaderText = "Email Address", Order = 3)]
        public required string Email { get; init; }

        [Column(HeaderText = "Hire Date", Order = 4)]
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