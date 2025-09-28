using ExcelsiorClosedXml;
using Microsoft.EntityFrameworkCore;

public class BindingModel
{
    #region DataModel

    public class Address
    {
        public required int StreetNumber { get; init; }

        public required string Street { get; init; }
    }

    public class Company
    {
        public required int Id { get; init; }

        public required string Name { get; init; }
    }

    public class Employee
    {
        public required int Id { get; init; }

        public required string Name { get; init; }

        public required Company Company { get; init; }

        public required Address Address { get; init; }

        public required string Email { get; init; }
    }

    #endregion

    #region EmployeeBindingModel

    public class EmployeeBindingModel
    {
        public required string Name { get; init; }

        public required string Email { get; init; }

        public required string Company { get; init; }

        public required string Address { get; init; }
    }

    #endregion

    public class TheDbContext :
        DbContext
    {
        public DbSet<Company> Companies { get; set; } = null!;
        public DbSet<Employee> Employees { get; set; } = null!;
        public DbSet<Address> Addresses { get; set; } = null!;
    }

    static void Foo(TheDbContext dbContext)
    {
        #region ModelProjection

        var employees = dbContext
            .Employees
            .Select(_ =>
                new EmployeeBindingModel
                {
                    Name = _.Name,
                    Email = _.Email,
                    Company = _.Company.Name,
                    Address = $"{_.Address.StreetNumber} {_.Address.Street}",
                });
        var builder = new BookBuilder();
        builder.AddSheet(employees);

        #endregion
    }
}