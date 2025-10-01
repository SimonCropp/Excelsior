using Excelsior;

public class Employee
{
    [Column(Heading = "Employee ID", Order = 1)]
    public required int Id { get; init; }

    [Column(Heading = "Full Name", Order = 2)]
    public required string Name { get; init; }

    [Column(Heading = "Email Address", Order = 3)]
    public required string Email { get; init; }

    [Column(Heading = "Hire Date", Order = 4)]
    public Date? HireDate { get; init; }

    [Column(Heading = "Annual Salary", Order = 5)]
    public int Salary { get; init; }

    public bool IsActive { get; init; }

    public EmployeeStatus Status { get; init; }
}