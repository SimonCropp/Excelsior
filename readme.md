# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/2t806jcx34s3r796/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/Replicant)
[![NuGet Status](https://img.shields.io/nuget/v/Excelsior.svg)](https://www.nuget.org/packages/Excelsior/)

**See [Milestones](../../milestones?state=closed) for release notes.**


## NuGet package

https://nuget.org/packages/Excelsior/


## Usage


### Model

Given an input class:

<!-- snippet: Employee.cs -->
<a id='snippet-Employee.cs'></a>
```cs
public class Employee
{
    [Display(Name = "Employee ID", Order = 1)]
    public required int Id { get; set; }

    [Display(Name = "Full Name", Order = 2)]
    public required string Name { get; set; }

    [Display(Name = "Email Address", Order = 3)]
    public required string Email { get; set; }

    [Display(Name = "Hire Date", Order = 4)]
    public required DateTime HireDate { get; set; }

    [Display(Name = "Annual Salary", Order = 5)]
    public required decimal Salary { get; set; }

    [Display(Name = "Is Active", Order = 7)]
    public required bool IsActive { get; set; }

    public required EmployeeStatus Status { get; set; }
}
```
<sup><a href='/src/Tests/Sample/Employee.cs#L1-L22' title='Snippet source file'>snippet source</a> | <a href='#snippet-Employee.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Builder

<!-- snippet: Usage -->
<a id='snippet-Usage'></a>
```cs
List<Employee> employees =
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
builder.AddSheet(employees);

var book = builder.Build();
```
<sup><a href='/src/Tests/Tests.cs#L7-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/simple.png">
