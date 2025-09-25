public class Employee
{
    [Display(Name = "Employee ID", Order = 1)]
    public required int Id { get; init; }

    [Display(Name = "Full Name", Order = 2)]
    public required string Name { get; init; }

    [Display(Name = "Email Address", Order = 3)]
    public required string Email { get; init; }

    [Display(Name = "Hire Date", Order = 4)]
    public DateTime HireDate { get; init; }

    [Display(Name = "Annual Salary", Order = 5)]
    public decimal Salary { get; init; }

    public bool IsActive { get; init; }

    public  EmployeeStatus Status { get; init; }
}