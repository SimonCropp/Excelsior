# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/fo33wu7ud6es1t2o/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/Excelsior)
[![NuGet Status](https://img.shields.io/nuget/v/Excelsior.svg)](https://www.nuget.org/packages/Excelsior/)

A data driven generator for Excel spreadsheets. Leverages [ClosedXML](https://github.com/ClosedXML/ClosedXML).

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


### Saving to a stream

<!-- snippet: ToStream -->
<a id='snippet-ToStream'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data);

var stream = new MemoryStream();
builder.ToStream(stream);
```
<sup><a href='/src/Tests/Tests.cs#L368-L376' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### CustomHeaders

<!-- snippet: CustomHeaders -->
<a id='snippet-CustomHeaders'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(
        _ => _.Name,
        _ => _.HeaderText = "Employee Name")
    .Column(
        _ => _.Email,
        _ => _.HeaderText = "Email Address");
```
<sup><a href='/src/Tests/Tests.cs#L48-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeaders' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnOrdering

<!-- snippet: ColumnOrdering -->
<a id='snippet-ColumnOrdering'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(_ => _.Email, _ => _.Order = 1)
    .Column(_ => _.Name, _ => _.Order = 2)
    .Column(_ => _.Salary, _ => _.Order = 3);
```
<sup><a href='/src/Tests/Tests.cs#L71-L79' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Tests.cs#L91-L102' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeaderStyle' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Tests.cs#L114-L125' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ConditionalStyling

<!-- snippet: ConditionalStyling -->
<a id='snippet-ConditionalStyling'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Salary,
        config =>
        {
            config.ConditionalStyling = (style, value) =>
            {
                if (value > 100000)
                {
                    style.Font.FontColor = XLColor.DarkGreen;
                    style.Font.Bold = true;
                }
            };
        })
    .Column(
        _ => _.IsActive,
        config =>
        {
            config.ConditionalStyling = (style, isActive) =>
            {
                var fill = style.Fill;
                if (isActive)
                {
                    fill.BackgroundColor = XLColor.LightGreen;
                }
                else
                {
                    fill.BackgroundColor = XLColor.LightPink;
                }
            };
        });
```
<sup><a href='/src/Tests/Tests.cs#L137-L172' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConditionalStyling' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Render

<!-- snippet: CustomRender -->
<a id='snippet-CustomRender'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(
        _ => _.Email,
        _ => _.Render = value => $"📧 {value}")
    .Column(
        _ => _.IsActive,
        _ => _.Render = active => active ? "✓ Active" : "✗ Inactive")
    .Column(
        _ => _.HireDate,
        _ => _.Format = "yyyy-MM-dd");
```
<sup><a href='/src/Tests/Tests.cs#L184-L198' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### WorksheetName

<!-- snippet: WorksheetName -->
<a id='snippet-WorksheetName'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, "Employee Report");
```
<sup><a href='/src/Tests/Tests.cs#L210-L215' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnWidths

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(_ => _.Name, _ => _.ColumnWidth = 25)
    .Column(_ => _.Email, _ => _.ColumnWidth = 30)
    .Column(_ => _.HireDate, _ => _.ColumnWidth = 15);
```
<sup><a href='/src/Tests/Tests.cs#L227-L235' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->



### Complex Types

For complex types, by default is to render via `.ToString()`.


#### Models

<!-- snippet: ComplexTypeModels -->
<a id='snippet-ComplexTypeModels'></a>
```cs
public record Person(string Name, Address Address);

public record Address(int Number, string Street, State State, string City, ushort PostCode);
```
<sup><a href='/src/Tests/ComplexTypeWithToString.cs#L9-L15' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeModels' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Build

<!-- snippet: ComplexTypeWithToString -->
<a id='snippet-ComplexTypeWithToString'></a>
```cs
List<Person> data =
[
    new("John Doe",
        new Address(
            Number: 900,
            Street: "Victoria Square",
            State: State.SouthAustralia,
            City: "Adelaide",
            PostCode: 5000)),
];

var builder = new BookBuilder();
builder.AddSheet(data);

var book = builder.Build();
```
<sup><a href='/src/Tests/ComplexTypeWithToString.cs#L20-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: ComplexTypeWithToString.Test.verified.csv -->
<a id='snippet-ComplexTypeWithToString.Test.verified.csv'></a>
```csv
Name,Address
John Doe,"Address { Number = 900, Street = Victoria Square, State = SouthAustralia, City = Adelaide, PostCode = 5000 }"
```
<sup><a href='/src/Tests/ComplexTypeWithToString.Test.verified.csv#L1-L2' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString.Test.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Custom render for Complex Types

<!-- snippet: ComplexTypeWithCustomRender -->
<a id='snippet-ComplexTypeWithCustomRender'></a>
```cs
List<Person> data =
[
    new("John Doe",
        new Address(
            Number: 900,
            Street: "Victoria Square",
            State: State.SouthAustralia,
            City: "Adelaide",
            PostCode: 5000)),
];

BookBuilder.RenderFor<Address>(
    _ => $"{_.Number}, {_.Street}, {_.State}, {_.City}, {_.PostCode}");

var builder = new BookBuilder();
builder.AddSheet(data);

var book = builder.Build();
```
<sup><a href='/src/Tests/ComplexTypeWithCustomRender.cs#L16-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Icon

[Excel](https://thenounproject.com/icon/excel-4558727/) designed by [Start Up Graphic Design](https://thenounproject.com/creator/ppanggm/) from [The Noun Project](https://thenounproject.com/).
