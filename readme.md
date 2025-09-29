# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/fo33wu7ud6es1t2o/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/Excelsior)
[![NuGet Status](https://img.shields.io/nuget/v/ExcelsiorClosedXml.svg?label=ExcelsiorClosedXml)](https://www.nuget.org/packages/ExcelsiorClosedXml/)
[![NuGet Status](https://img.shields.io/nuget/v/ExcelsiorClosedXml.svg?label=ExcelsiorAspose)](https://www.nuget.org/packages/ExcelsiorAspose/)

Excelsior is an Excel spreadsheet generation library with a distinctive data-driven approach.

**See [Milestones](../../milestones?state=closed) for release notes.**


## Supported libraries

The architecture is designed to support multiple spreadsheet creation libraries.

Currently supported libraries include:

 * [ClosedXML](https://github.com/ClosedXML/ClosedXML) via the [ExcelsiorClosedXml](https://nuget.org/packages/ExcelsiorClosedXml/) nuget
 * [Aspose.Cells](https://docs.aspose.com/cells/net/) via the [ExcelsiorAspose](https://nuget.org/packages/ExcelsiorAspose/) nuget


## Usage


### Model

Given an input class:

<!-- snippet: Employee.cs -->
<a id='snippet-Employee.cs'></a>
```cs
public class Employee
{
    [Display(Name = "Employee ID", Order = 1)]
    public required int Id { get; init; }

    [Display(Name = "Full Name", Order = 2)]
    public required string Name { get; init; }

    [Display(Name = "Email Address", Order = 3)]
    public required string Email { get; init; }

    [Display(Name = "Hire Date", Order = 4)]
    public DateTime? HireDate { get; init; }

    [Display(Name = "Annual Salary", Order = 5)]
    public int Salary { get; init; }

    public bool IsActive { get; init; }

    public EmployeeStatus Status { get; init; }
}
```
<sup><a href='/src/Model/Employee.cs#L1-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-Employee.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`[DisplayAttribute]` is optional. If omitted:

 * Order is based on the order of the properties defined in the class. Order can be [programmatically controlled](#column-ordering)
 * Header text is based on the property names that is camel case split. Headers can be [programmatically controlled](#custom-headers)


### Builder

`BookBuilder` is the root entry point.

Once instantiated the data for multiple sheets can be added.

<!-- snippet: Usage -->
<a id='snippet-Usage'></a>
```cs
var builder = new BookBuilder();

List<Employee> data =
[
    new()
    {
        Id = 1,
        Name = "John Doe",
        Email = "john@company.com",
        HireDate = new(2020, 1, 15),
        Salary = 75000,
        IsActive = true,
        Status = EmployeeStatus.FullTime
    },
    new()
    {
        Id = 2,
        Name = "Jane Smith",
        Email = "jane@company.com",
        HireDate = new(2019, 3, 22),
        Salary = 120000,
        IsActive = true,
        Status = EmployeeStatus.FullTime
    },
];
builder.AddSheet(data);

var book = await builder.Build();
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L40-L71' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.Simple_Sheet1.png">


### Saving to a stream

The above sample builds a instance of the Workbook for the target library. `Aspose.Cells.Workbook` for Aspose, and `ClosedXML.Excel.IXLWorkbook` for ClosedXml.

To instead save to a stream use `ToStream()`.

<!-- snippet: ToStream -->
<a id='snippet-ToStream'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data);

var stream = new MemoryStream();
await builder.ToStream(stream);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L636-L644' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Custom Headers

The header text for a column can be overridden:

<!-- snippet: CustomHeaders -->
<a id='snippet-CustomHeaders'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Name,
        _ => _.HeaderText = "Employee Name");
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L265-L273' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeaders' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.CustomHeaders_Sheet1.png">


### Column Ordering

The column order can be overridden:

<!-- snippet: ColumnOrdering -->
<a id='snippet-ColumnOrdering'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Email, _ => _.Order = 1)
    .Column(_ => _.Name, _ => _.Order = 2)
    .Column(_ => _.Salary, _ => _.Order = 3);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L285-L293' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.ColumnOrdering_Sheet1.png">


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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L305-L316' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeaderStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.HeaderStyle_Sheet1.png">


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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L328-L339' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.GlobalStyle_Sheet1.png">


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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L351-L386' title='Snippet source file'>snippet source</a> | <a href='#snippet-ConditionalStyling' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.ConditionalStyling_Sheet1.png">


### Render

<!-- snippet: CustomRender -->
<a id='snippet-CustomRender'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Name,
        _ => _.Render = value => value.ToUpper())
    .Column(
        _ => _.IsActive,
        _ => _.Render = active => active ? "✓ Active" : "✗ Inactive")
    .Column(
        _ => _.HireDate,
        _ => _.Format = "yyyy-MM-dd");
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L398-L412' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.Render_Sheet1.png">


### Worksheet Name

<!-- snippet: WorksheetName -->
<a id='snippet-WorksheetName'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, "Employee Report");
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L424-L429' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Column Widths

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Name, _ => _.ColumnWidth = 25)
    .Column(_ => _.Email, _ => _.ColumnWidth = 30)
    .Column(_ => _.HireDate, _ => _.ColumnWidth = 15);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L441-L449' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.ColumnWidths_Sheet1.png">


### Complex Types

For complex types, by default is to render via `.ToString()`.


#### Models

<!-- snippet: ComplexTypeModels -->
<a id='snippet-ComplexTypeModels'></a>
```cs
public record Person(string Name, Address Address);

public record Address(int Number, string Street, string City, State State, ushort PostCode);
```
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithToString.cs#L9-L15' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeModels' title='Start of snippet'>anchor</a></sup>
<a id='snippet-ComplexTypeModels-1'></a>
```cs
public record Person(string Name, Address Address);

public record Address(int Number, string Street, string City, State State, ushort PostCode);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/ComplexTypeWithToString.cs#L10-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeModels-1' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Build

<!-- snippet: ComplexTypeWithToString -->
<a id='snippet-ComplexTypeWithToString'></a>
```cs
var builder = new BookBuilder();

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

builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithToString.cs#L21-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString' title='Start of snippet'>anchor</a></sup>
<a id='snippet-ComplexTypeWithToString-1'></a>
```cs
var builder = new BookBuilder();

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
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/ComplexTypeWithToString.cs#L21-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString-1' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/ComplexTypeWithToString.Test_Sheet1.png">


### Custom render for Complex Types

<!-- snippet: ComplexTypeWithCustomRender -->
<a id='snippet-ComplexTypeWithCustomRender'></a>
```cs
var builder = new BookBuilder();

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

ValueRenderer.For<Address>(
    _ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithCustomRender.cs#L18-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender' title='Start of snippet'>anchor</a></sup>
<a id='snippet-ComplexTypeWithCustomRender-1'></a>
```cs
ValueRenderer.For<Address>(
    _ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");

var builder = new BookBuilder();

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
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/ComplexTypeWithCustomRender.cs#L17-L36' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender-1' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/ComplexTypeWithCustomRender.Test_Sheet1.png">


### Whitespace

By default whitespace is trimmed

<!-- snippet: Whitespace -->
<a id='snippet-Whitespace'></a>
```cs
var builder = new BookBuilder();

List<Employee> data =
[
    new()
    {
        Id = 1,
        Name = "    John Doe   ",
        Email = "    john@company.com    ",
    }
];
builder.AddSheet(data);

var book = await builder.Build();
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L78-L95' title='Snippet source file'>snippet source</a> | <a href='#snippet-Whitespace' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/Tests.Whitespace_Sheet1.png">


#### Disable whitespace trimming

<!-- snippet: DisableWhitespaceTrim -->
<a id='snippet-DisableWhitespaceTrim'></a>
```cs
List<Employee> data =
[
    new()
    {
        Id = 1,
        Name = "    John Doe   ",
        Email = "    john@company.com    ",
    }
];

var builder = new BookBuilder(trimWhitespace: false);
builder.AddSheet(data);

var book = await builder.Build();
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L102-L119' title='Snippet source file'>snippet source</a> | <a href='#snippet-DisableWhitespaceTrim' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/Tests.DisableWhitespaceTrim_Sheet1.png">


### Enumerable string properties

Properties that are castable to an `IEnumerable<string>` will automatically be rendered as a point form list.


#### Module

<!-- snippet: EnumerableModel -->
<a id='snippet-EnumerableModel'></a>
```cs
public record Person(string Name, string[] PhoneNumbers);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/EnumerableStringTests.cs#L4-L8' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Render

<!-- snippet: EnumerableUsage -->
<a id='snippet-EnumerableUsage'></a>
```cs
List<Person> data =
[
    new("John Doe",
        PhoneNumbers:
        [
            "+1 3057380950",
            "+1 5056169368",
            "+1 8634446859"
        ]),
];

var builder = new BookBuilder();
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/EnumerableStringTests.cs#L13-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/EnumerableStringTests.Test_Sheet1.png">


## Binding Model

The recommended approach is to use are specific type for binding.

This will make configuration and rendering simpler. It will often also result in better performance. The reason being that the projection into the binding type can be done by the database via an ORM. This will result in a faster query response and less data being transferred from the database.

Take for example of rendering employees to a sheet. A potential model could be `Company`, `Employee`, and `Address`.

<!-- snippet: DataModel -->
<a id='snippet-DataModel'></a>
```cs
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
```
<sup><a href='/src/Model/BindingModel.cs#L6-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-DataModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Then a custom binding type can be used.

<!-- snippet: EmployeeBindingModel -->
<a id='snippet-EmployeeBindingModel'></a>
```cs
public class EmployeeBindingModel
{
    public required string Name { get; init; }
    public required string Email { get; init; }
    public required string Company { get; init; }
    public required string Address { get; init; }
}
```
<sup><a href='/src/Model/BindingModel.cs#L31-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-EmployeeBindingModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The custom binding type can be queried and  rendered into a sheet.

<!-- snippet: ModelProjection -->
<a id='snippet-ModelProjection'></a>
```cs
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
```
<sup><a href='/src/Model/BindingModel.cs#L53-L68' title='Snippet source file'>snippet source</a> | <a href='#snippet-ModelProjection' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Icon

[Excel](https://thenounproject.com/icon/excel-4558727/) designed by [Start Up Graphic Design](https://thenounproject.com/creator/ppanggm/) from [The Noun Project](https://thenounproject.com/).
