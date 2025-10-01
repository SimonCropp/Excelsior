# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/fo33wu7ud6es1t2o/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/Excelsior)
[![NuGet Status](https://img.shields.io/nuget/v/ExcelsiorClosedXml.svg?label=ExcelsiorClosedXml)](https://www.nuget.org/packages/ExcelsiorClosedXml/)
[![NuGet Status](https://img.shields.io/nuget/v/ExcelsiorClosedXml.svg?label=ExcelsiorAspose)](https://www.nuget.org/packages/ExcelsiorAspose/)

Excelsior is a Excel spreadsheet generation library with a distinctive data-driven approach.

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
```
<sup><a href='/src/Model/Employee.cs#L1-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-Employee.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`[ColumnAttribute]` is optional. If omitted:

 * Order is based on the order of the properties defined in the class. Order can be [programmatically controlled](#column-ordering)
 * Heading text is based on the property names that is camel case split. Headings can be [programmatically controlled](#custom-headings)


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
<sup><a href='/src/ExcelsiorAspose.Tests/UsageTests.cs#L7-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.Simple_Sheet1.png">


### DataAnnotations Attributes

`DisplayAttribute` and `DisplayNameAttribute` from System.ComponentModel.DataAnnotations are supported.

<!-- snippet: DataAnnotationsModel -->
<a id='snippet-DataAnnotationsModel'></a>
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
    public Date? HireDate { get; init; }

    [Display(Name = "Annual Salary", Order = 5)]
    public int Salary { get; init; }

    [DisplayName("IsActive")]
    public bool IsActive { get; init; }

    public EmployeeStatus Status { get; init; }
}

public enum EmployeeStatus
{
    [Display(Name = "Full Time")]
    FullTime,

    [Display(Name = "Part Time")]
    PartTime,

    [Display(Name = "Contract")]
    Contract,

    [Display(Name = "Terminated")]
    Terminated
}
```
<sup><a href='/src/ExcelsiorAspose.Tests/DataAnnotationsTests.cs#L43-L83' title='Snippet source file'>snippet source</a> | <a href='#snippet-DataAnnotationsModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L248-L256' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Custom Headings

The heading text for a column can be overridden:

<!-- snippet: CustomHeadings -->
<a id='snippet-CustomHeadings'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Name,
        _ => _.Heading = "Employee Name");
```
<sup><a href='/src/ExcelsiorAspose.Tests/Headings.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeadings' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.CustomHeadings_Sheet1.png">


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
<sup><a href='/src/ExcelsiorAspose.Tests/ColumnOrdering.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.ColumnOrdering_Sheet1.png">


### Heading Style

<!-- snippet: HeadingStyle -->
<a id='snippet-HeadingStyle'></a>
```cs
var builder = new BookBuilder(
    headingStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.FontColor = XLColor.White;
        style.Fill.BackgroundColor = XLColor.DarkBlue;
    });
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L10-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeadingStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.HeadingStyle_Sheet1.png">


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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L33-L44' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.GlobalStyle_Sheet1.png">


### Cell Styling

<!-- snippet: CellStyle -->
<a id='snippet-CellStyle'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Salary,
        config =>
        {
            config.CellStyle = (style, employee, salary) =>
            {
                if (salary > 100000)
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
            config.CellStyle = (style, employee, isActive) =>
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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L56-L91' title='Snippet source file'>snippet source</a> | <a href='#snippet-CellStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/Tests.CellStyle_Sheet1.png">


### Render

<!-- snippet: CustomRender -->
<a id='snippet-CustomRender'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Name,
        _ => _.Render = (employee, name) => name.ToUpper())
    .Column(
        _ => _.IsActive,
        _ => _.Render = (employee, active) => active ? "Active" : "Inactive")
    .Column(
        _ => _.HireDate,
        _ => _.Format = "yyyy-MM-dd");
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L103-L117' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomRender' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L129-L134' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Column Widths

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Name, _ => _.Width = 25)
    .Column(_ => _.Email, _ => _.Width = 30)
    .Column(_ => _.HireDate, _ => _.Width = 15);
```
<sup><a href='/src/ExcelsiorClosedXml.Tests/Tests.cs#L146-L154' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithToString.cs#L10-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeModels' title='Start of snippet'>anchor</a></sup>
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
<!-- endSnippet -->


#### Result:

<img src="/src/ExcelsiorClosedXml.Tests/ComplexTypeWithToString.Test_Sheet1.png">


### Custom render for Complex Types

<!-- snippet: ComplexTypeWithCustomRenderInit -->
<a id='snippet-ComplexTypeWithCustomRenderInit'></a>
```cs
[ModuleInitializer]
public static void Init() =>
    ValueRenderer.For<Address>(_ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");
```
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithCustomRender.cs#L14-L20' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRenderInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

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
builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorAspose.Tests/ComplexTypeWithCustomRender.cs#L25-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/ExcelsiorAspose.Tests/WhitespaceTests.cs#L7-L24' title='Snippet source file'>snippet source</a> | <a href='#snippet-Whitespace' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/WhitespaceTests.Whitespace_Sheet1.png">


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
<sup><a href='/src/ExcelsiorAspose.Tests/WhitespaceTests.cs#L32-L49' title='Snippet source file'>snippet source</a> | <a href='#snippet-DisableWhitespaceTrim' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/WhitespaceTests.Disable_Sheet1.png">


### Enumerable string properties

Properties that are castable to an `IEnumerable<string>` will automatically be rendered as a point form list.


#### Module

<!-- snippet: EnumerableModel -->
<a id='snippet-EnumerableModel'></a>
```cs
public record Person(string Name, string[] PhoneNumbers);
```
<sup><a href='/src/ExcelsiorAspose.Tests/EnumerableStringTests.cs#L4-L8' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableModel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/ExcelsiorAspose.Tests/EnumerableStringTests.cs#L13-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableUsage' title='Start of snippet'>anchor</a></sup>
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


### ColumnAttribute

`ColumnAttribute` allows customization of rendering at the model level.

It is intended as an alternative to the usage of `DisplayAttribute` and `DisplayNameAttribute`.
`DisplayAttribute` and `DisplayNameAttribute` are support for scenarios where it is not convenient to reference Excelsior from that assembly.


#### ColumnAttribute definition

<!-- snippet: ColumnAttribute.cs -->
<a id='snippet-ColumnAttribute.cs'></a>
```cs
namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class ColumnAttribute :
    Attribute
{
    public string? Heading { get; set; }
    public int Order { get; set; } = -1;
    public double Width { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public bool IsHtml { get; set; }
}
```
<sup><a href='/src/Excelsior/ColumnAttribute.cs#L1-L13' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttribute.cs' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Usage

<!-- snippet: ColumnAttributeModel -->
<a id='snippet-ColumnAttributeModel'></a>
```cs
public class Employee
{
    [Column(Heading = "Employee ID", Order = 1, Format = "0000")]
    public required int Id { get; init; }

    [Column(Heading = "Full Name", Order = 2, Width = 20)]
    public required string Name { get; init; }

    [Column(Heading = "Email Address", Width = 30)]
    public required string Email { get; init; }

    [Column(Heading = "Hire Date", Order = 3, NullDisplay = "unknown")]
    public Date? HireDate { get; init; }
}
```
<sup><a href='/src/ExcelsiorAspose.Tests/ColumnAttributeTests.cs#L4-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttributeModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

<!-- snippet: ColumnAttribute -->
<a id='snippet-ColumnAttribute'></a>
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
    },
    new()
    {
        Id = 2,
        Name = "Jane Smith",
        Email = "jane@company.com",
        HireDate = null,
    }
];

builder.AddSheet(data);
```
<sup><a href='/src/ExcelsiorAspose.Tests/ColumnAttributeTests.cs#L26-L50' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttribute' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/ExcelsiorClosedXml.Tests/ColumnAttributeTests.Test_Sheet1.png">


## Icon

[Grim Fandango](https://github.com/PapirusDevelopmentTeam/papirus-icon-theme/blob/master/Papirus/64x64/apps/grim-fandango-remastered.svg) from [Papirus Icons](https://github.com/PapirusDevelopmentTeam/papirus-icon-theme).

The [Excelsior Line](https://grim-fandango.fandom.com/wiki/Excelsior_Line) is a travel package sold by Manuel Calavera in the Lucas Arts game "Grim Fandango". The package consists of nothing more than a walking stick with a compass in the handle.
