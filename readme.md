# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/fo33wu7ud6es1t2o/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/Excelsior)
[![NuGet Status](https://img.shields.io/nuget/v/Excelsior.svg)](https://www.nuget.org/packages/Excelsior/)

Excelsior is an Excel spreadsheet generation library with a distinctive data-driven approach that leverages [ClosedXML](https://github.com/ClosedXML/ClosedXML) to create Excel-compatible files.

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

    public required bool IsActive { get; set; }

    public required EmployeeStatus Status { get; set; }
}
```
<sup><a href='/src/Tests/Sample/Employee.cs#L1-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-Employee.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`[DisplayAttribute]` is optional. If it is omitted:

 * Order is based on the order of the properties defined in the class. Order can be [programmatically controlled](#column-ordering)
 * Header text is based on the property names that is camel case split. Headers can be [programmatically controlled](#custom-headers)


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

var book = await builder.Build();
```
<sup><a href='/src/Tests/Tests.cs#L42-L73' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/Tests/Tests.Simple.DotNet_Sheet1.png">


### Saving to a stream

<!-- snippet: ToStream -->
<a id='snippet-ToStream'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data);

var stream = new MemoryStream();
await builder.ToStream(stream);
```
<sup><a href='/src/Tests/Tests.cs#L522-L530' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Custom Headers

<!-- snippet: CustomHeaders -->
<a id='snippet-CustomHeaders'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(
        _ => _.Name,
        _ => _.HeaderText = "Employee Name");
```
<sup><a href='/src/Tests/Tests.cs#L205-L213' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeaders' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.CustomHeaders.DotNet_Sheet1.png">


### Column Ordering

<!-- snippet: ColumnOrdering -->
<a id='snippet-ColumnOrdering'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(_ => _.Email, _ => _.Order = 1)
    .Column(_ => _.Name, _ => _.Order = 2)
    .Column(_ => _.Salary, _ => _.Order = 3);
```
<sup><a href='/src/Tests/Tests.cs#L225-L233' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.ColumnOrdering.DotNet_Sheet1.png">


### Header Style

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
<sup><a href='/src/Tests/Tests.cs#L245-L256' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeaderStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.HeaderStyle.DotNet_Sheet1.png">


### Global Style

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
<sup><a href='/src/Tests/Tests.cs#L268-L279' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.GlobalStyle.DotNet_Sheet1.png">


### Conditional Styling

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
<sup><a href='/src/Tests/Tests.cs#L291-L326' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConditionalStyling' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.ConditionalStyling.DotNet_Sheet1.png">


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
<sup><a href='/src/Tests/Tests.cs#L338-L352' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.Render.DotNet_Sheet1.png">


### Worksheet Name

<!-- snippet: WorksheetName -->
<a id='snippet-WorksheetName'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, "Employee Report");
```
<sup><a href='/src/Tests/Tests.cs#L364-L369' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Column Widths

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data)
    .Column(_ => _.Name, _ => _.ColumnWidth = 25)
    .Column(_ => _.Email, _ => _.ColumnWidth = 30)
    .Column(_ => _.HireDate, _ => _.ColumnWidth = 15);
```
<sup><a href='/src/Tests/Tests.cs#L381-L389' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/Tests.ColumnWidths.DotNet_Sheet1.png">


### Complex Types

For complex types, by default is to render via `.ToString()`.


#### Models

<!-- snippet: ComplexTypeModels -->
<a id='snippet-ComplexTypeModels'></a>
```cs
public record Person(string Name, Address Address);

public record Address(int Number, string Street, string City, State State, ushort PostCode);
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
            City: "Adelaide",
            State: State.SA,
            PostCode: 5000)),
];

var builder = new BookBuilder();
builder.AddSheet(data);
```
<sup><a href='/src/Tests/ComplexTypeWithToString.cs#L21-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/ComplexTypeWithToString.Test.DotNet_Sheet1.png">


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
            City: "Adelaide",
            State: State.SA,
            PostCode: 5000)),
];

BookBuilder.RenderFor<Address>(
    _ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");

var builder = new BookBuilder();
builder.AddSheet(data);
```
<sup><a href='/src/Tests/ComplexTypeWithCustomRender.cs#L17-L36' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Tests/ComplexTypeWithCustomRender.Test.DotNet_Sheet1.png">


## Icon

[Excel](https://thenounproject.com/icon/excel-4558727/) designed by [Start Up Graphic Design](https://thenounproject.com/creator/ppanggm/) from [The Noun Project](https://thenounproject.com/).
