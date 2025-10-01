using Excelsior;

public class Employee
{
    [Column(Header = "Employee ID", Order = 1)]
    public required int Id { get; init; }

    [Column(Header = "Full Name", Order = 2)]
    public required string Name { get; init; }

    [Column(Header = "Email Address", Order = 3)]
    public required string Email { get; init; }

    [Column(Header = "Hire Date", Order = 4)]
    public Date? HireDate { get; init; }

    [Column(Header = "Annual Salary", Order = 5)]
    public int Salary { get; init; }

    public bool IsActive { get; init; }

    public EmployeeStatus Status { get; init; }
}