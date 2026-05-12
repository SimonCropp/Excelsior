using Excelsior;

public class Employee
{
    [Column(Heading = "Employee ID", Order = 1)]
    public required int Id;

    [Column(Heading = "Full Name", Order = 2)]
    public required string Name;

    [Column(Heading = "Email Address", Order = 3)]
    public required string Email;

    [Column(Heading = "Hire Date", Order = 4)]
    public Date? HireDate;

    [Column(Heading = "Annual Salary", Order = 5)]
    public int Salary;

    public bool IsActive;

    public EmployeeStatus Status;
}