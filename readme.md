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
List<Employee> data =
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
builder.AddSheet(data);

var book = builder.Build();
```
<sup><a href='/src/Tests/Tests.cs#L7-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/simple.png">


### CustomHeaders

<!-- snippet: CustomHeaders -->
<a id='snippet-CustomHeaders'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .ConfigureColumn(
        nameof(Employee.Name),
        _ => _.HeaderText = "Employee Name")
    .ConfigureColumn(
        nameof(Employee.Email),
        _ => _.HeaderText = "Email Address");

var book = builder.Build();
```
<sup><a href='/src/Tests/Tests.cs#L47-L60' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeaders' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnOrdering

<!-- snippet: ColumnOrdering -->
<a id='snippet-ColumnOrdering'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .ConfigureColumn(nameof(Employee.Email), _ => _.Order = 1)
    .ConfigureColumn(nameof(Employee.Name), _ => _.Order = 2)
    .ConfigureColumn(nameof(Employee.Salary), _ => _.Order = 3);
```
<sup><a href='/src/Tests/Tests.cs#L69-L76' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnOrdering

<!-- snippet: ColumnOrdering -->
<a id='snippet-ColumnOrdering'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .ConfigureColumn(nameof(Employee.Email), _ => _.Order = 1)
    .ConfigureColumn(nameof(Employee.Name), _ => _.Order = 2)
    .ConfigureColumn(nameof(Employee.Salary), _ => _.Order = 3);
```
<sup><a href='/src/Tests/Tests.cs#L69-L76' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### HeaderStyle

<!-- snippet: HeaderStyle -->
<a id='snippet-HeaderStyle'></a>
```cs
var builder = new BookBuilder(
    headerStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.FontColor = XLColor.White;
        style.Fill.BackgroundColor = XLColor.DarkBlue;
    });
builder.AddSheet(data);
```
<sup><a href='/src/Tests/Tests.cs#L86-L95' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeaderStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### GlobalStyle

<!-- snippet: GlobalStyle -->
<a id='snippet-GlobalStyle'></a>
```cs
var builder = new BookBuilder(
    globalStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.FontColor = XLColor.White;
        style.Fill.BackgroundColor = XLColor.DarkBlue;
    });
builder.AddSheet(data);
```
<sup><a href='/src/Tests/Tests.cs#L105-L114' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ConditionalStyling

<!-- snippet: ConditionalStyling -->
<a id='snippet-ConditionalStyling'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .ConfigureColumn(nameof(Employee.Salary), config =>
    {
        config.ConditionalStyling = (style, value) =>
        {
            if (value is decimal and > 100000)
            {
                style.Font.FontColor = XLColor.DarkGreen;
                style.Font.Bold = true;
            }
        };
    })
    .ConfigureColumn(nameof(Employee.IsActive), config =>
    {
        config.ConditionalStyling = (style, value) =>
        {
            if (value is bool isActive)
            {
                style.Fill.BackgroundColor = isActive ? XLColor.LightGreen : XLColor.LightPink;
            }
        };
    });
```
<sup><a href='/src/Tests/Tests.cs#L125-L149' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConditionalStyling' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### CustomFormatters

<!-- snippet: CustomFormatters -->
<a id='snippet-CustomFormatters'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .ConfigureColumn(
        nameof(Employee.Email),
        _ => _.CustomFormatter = value => $"📧 {value}")
    .ConfigureColumn(
        nameof(Employee.IsActive),
        _ => _.BooleanDisplayFormat = active => active ? "✓ Active" : "✗ Inactive")
    .ConfigureColumn(
        nameof(Employee.HireDate),
        _ => _.DateTimeFormat = "yyyy-MM-dd");
```
<sup><a href='/src/Tests/Tests.cs#L160-L173' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomFormatters' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### WorksheetName

<!-- snippet: WorksheetName -->
<a id='snippet-WorksheetName'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, "Employee Report");
```
<sup><a href='/src/Tests/Tests.cs#L183-L187' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnWidths

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .ConfigureColumn(nameof(Employee.Name), _ => _.ColumnWidth = 25)
    .ConfigureColumn(nameof(Employee.Email), _ => _.ColumnWidth = 30)
    .ConfigureColumn(nameof(Employee.HireDate), _ => _.ColumnWidth = 15);
```
<sup><a href='/src/Tests/Tests.cs#L197-L204' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->
