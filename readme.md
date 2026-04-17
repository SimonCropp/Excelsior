# <img src="/src/icon.png" height="30px"> Excelsior

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/Excelsior)](https://ci.appveyor.com/project/SimonCropp/Excelsior)
[![NuGet Status](https://img.shields.io/nuget/v/Excelsior.svg?label=Excelsior)](https://www.nuget.org/packages/Excelsior/)

Excelsior is a Excel spreadsheet generation library with a distinctive data-driven approach. Uses [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) for spreadsheet creation and [OpenXmlHtml](https://github.com/SimonCropp/OpenXmlHtml) for HTML cell rendering.

**See [Milestones](../../milestones?state=closed) for release notes.**


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

Once instantiated, the data for multiple sheets can be added.

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

using var book = await builder.Build();
```
<sup><a href='/src/Excelsior.Tests/UsageTests.cs#L7-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-Usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Result:

<img src="/src/Excelsior.Tests/UsageTests.Test_Sheet1.png">


### Worksheet Name

Worksheet defaults to `SheetN`, when `N` is a counter. So the first sheet is `Sheet1`, the second is `Sheet2`, etc.

The name can be controlled by passing an explicit value.

<!-- snippet: WorksheetName -->
<a id='snippet-WorksheetName'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, "Employee Report");
```
<sup><a href='/src/Excelsior.Tests/WorksheetName.cs#L9-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-WorksheetName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### DataAnnotations Attributes

`DisplayAttribute` and `DisplayNameAttribute` from [System.ComponentModel.DataAnnotations](https://www.nuget.org/packages/system.componentmodel.annotations/) are supported.

`DisplayAttribute` and `DisplayNameAttribute` are support for scenarios where it is not convenient to reference Excelsior from that assembly.<!-- singleLineInclude: DisplayAttributeScenario. path: /docs/DisplayAttributeScenario.include.md -->

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
<sup><a href='/src/Excelsior.Tests/DataAnnotationsTests.cs#L43-L83' title='Snippet source file'>snippet source</a> | <a href='#snippet-DataAnnotationsModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/DataAnnotationsTests.Simple_Sheet1.png">


### Saving to a stream

To save to a stream use `ToStream()`.

<!-- snippet: ToStream -->
<a id='snippet-ToStream'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(data);

var stream = new MemoryStream();
await builder.ToStream(stream);
```
<sup><a href='/src/Excelsior.Tests/Saving.cs#L10-L18' title='Snippet source file'>snippet source</a> | <a href='#snippet-ToStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Grouping Column Configuration

When applying multiple settings to the same column, prefer grouping them in a single `Column` call rather than using separate method calls. This makes it clearer which settings belong together.

```cs
// prefer
sheet.Column(
    _ => _.Name,
    _ =>
    {
        _.Width = 25;
        _.IsHtml = true;
        _.Render = (row, _) => $"<a href='/people/{row.Id}'>{row.Name}</a>";
    });

// over
sheet.Width(_ => _.Name, 25);
sheet.IsHtml(_ => _.Name);
sheet.Render(_ => _.Name, (row, _) => $"<a href='/people/{row.Id}'>{row.Name}</a>");
```


### Custom Headings

The heading text for a column can be overridden:

#### Fluent

<!-- snippet: CustomHeadings -->
<a id='snippet-CustomHeadings'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Name,
        _ => _.Heading = "Employee Name");
```
<sup><a href='/src/Excelsior.Tests/Headings.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomHeadings' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### ColumnAttribute

```
public class Employee
{
    [Column(Heading = "Employee Name")]
    public required string Name { get; init; }
```


#### DisplayNameAttribute

```
public class Employee
{
    [DisplayName("Employee Name")]
    public required string Name { get; init; }
```


#### DisplayAttribute

```
public class Employee
{
    [Display(Name = "Employee Name")]
    public required string Name { get; init; }
```


#### Result:

<img src="/src/Excelsior.Tests/Headings.Fluent_Sheet1.png">


#### Order of precedence

 1. Fluent
 1. `ColumnAttribute`
 1. `DisplayAttribute`
 1. `DisplayNameAttribute`


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
<sup><a href='/src/Excelsior.Tests/ColumnOrdering.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnOrdering' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/ColumnOrdering.Fluent_Sheet1.png">


### Heading Style

<!-- snippet: HeadingStyle -->
<a id='snippet-HeadingStyle'></a>
```cs
var builder = new BookBuilder(
    headingStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.Color = "FFFFFF";
        style.BackgroundColor = "00008B";
    });
builder.AddSheet(data);
```
<sup><a href='/src/Excelsior.Tests/StyleTests.cs#L10-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-HeadingStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/StyleTests.HeadingStyle_Sheet1.png">


### Global Style

<!-- snippet: GlobalStyle -->
<a id='snippet-GlobalStyle'></a>
```cs
var builder = new BookBuilder(
    globalStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.Color = "FFFFFF";
        style.BackgroundColor = "00008B";
    });
builder.AddSheet(data);
```
<sup><a href='/src/Excelsior.Tests/StyleTests.cs#L33-L44' title='Snippet source file'>snippet source</a> | <a href='#snippet-GlobalStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/StyleTests.GlobalStyle_Sheet1.png">


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
                    style.Font.Color = "006400";
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
                if (isActive)
                {
                    style.BackgroundColor = "90EE90";
                }
                else
                {
                    style.BackgroundColor = "FFB6C1";
                }
            };
        });
```
<sup><a href='/src/Excelsior.Tests/StyleTests.cs#L56-L90' title='Snippet source file'>snippet source</a> | <a href='#snippet-CellStyle' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/StyleTests.CellStyle_Sheet1.png">


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
<sup><a href='/src/Excelsior.Tests/Render.cs#L10-L24' title='Snippet source file'>snippet source</a> | <a href='#snippet-CustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/Render.Fluent_Sheet1.png">


### Formula

A column can emit an Excel formula per row instead of a computed value. The
callback receives a `FormulaContext<TModel>` that exposes the current
1-based Excel `Row` number and helpers to build cell references to other
columns in the same row:

 * `Ref(_ => _.OtherProperty)` — full cell reference (e.g. `B5`).
 * `Column(_ => _.OtherProperty)` — column letter only (e.g. `B`).

Formulas take precedence over the normal value rendering, and may still use
`Format` for number formatting and `CellStyle` for styling.

<!-- snippet: FormulaFluent -->
<a id='snippet-FormulaFluent'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(
        _ => _.Salary,
        _ =>
        {
            _.Formula = (employee, context) =>
                $"={context.Ref(_ => _.Id)} * 10000";
            _.Format = "#,##0";
        });
```
<sup><a href='/src/Excelsior.Tests/FormulaTests.cs#L10-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-FormulaFluent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The shorter `Formula()` overload on `ISheetBuilder<TModel>` can be used when
the formula does not depend on the model:

```cs
builder.AddSheet(employees)
    .Formula(
        _ => _.Salary,
        context => $"={context.Ref(_ => _.Id)} * 1000");
```

#### Result:

<img src="/src/Excelsior.Tests/FormulaTests.Fluent_Sheet1.png">

**Note:** Formulas are an Excel-only feature. They are not supported in [Word tables](#word-tables) and will throw when `Build()` is called.


### Column Widths


#### Fluent

<!-- snippet: ColumnWidths -->
<a id='snippet-ColumnWidths'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Name, _ => _.Width = 25)
    .Column(_ => _.Email, _ => _.Width = 30)
    .Column(_ => _.HireDate, _ => _.Width = 15);
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnWidths' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### ColumnAttribute

```
public class Employee
{
    [Column(Width = 25)]
    public required string Name { get; init; }
```


#### Result:

<img src="/src/Excelsior.Tests/ColumnWidths.Fluent_Sheet1.png">


#### Order of precedence

 1. Fluent
 1. `ColumnAttribute`


### Column Min/Max Widths

Columns can be constrained to a minimum or maximum width while still auto-sizing based on content.

When `Width` is explicitly set, `MinWidth`/`MaxWidth` are ignored.

A book-wide or per-sheet `defaultMinColumnWidth` can also be set; it applies to every auto-sized column that does not have its own `MinWidth`. This pairs with the existing `defaultMaxColumnWidth` (default `50`).


#### Per-sheet default min width

<!-- snippet: SheetDefaultMinColumnWidth -->
<a id='snippet-SheetDefaultMinColumnWidth'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees, defaultMinColumnWidth: 25);
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L141-L146' title='Snippet source file'>snippet source</a> | <a href='#snippet-SheetDefaultMinColumnWidth' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Book-wide default min width

<!-- snippet: BookDefaultMinColumnWidth -->
<a id='snippet-BookDefaultMinColumnWidth'></a>
```cs
var builder = new BookBuilder(defaultMinColumnWidth: 25);
builder.AddSheet(employees);
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L158-L163' title='Snippet source file'>snippet source</a> | <a href='#snippet-BookDefaultMinColumnWidth' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### MinWidth

<!-- snippet: ColumnMinWidth -->
<a id='snippet-ColumnMinWidth'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Name, _ => _.MinWidth = 40);
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L57-L63' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnMinWidth' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### MaxWidth

<!-- snippet: ColumnMaxWidth -->
<a id='snippet-ColumnMaxWidth'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees)
    .Column(_ => _.Name, _ => _.MaxWidth = 5);
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L99-L105' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnMaxWidth' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Standalone methods

```cs
var sheet = builder.AddSheet(employees);
sheet.MinWidth(_ => _.Name, 40);
sheet.MaxWidth(_ => _.Email, 20);
```


#### ColumnAttribute

<!-- snippet: ColumnMinMaxWidthModel -->
<a id='snippet-ColumnMinMaxWidthModel'></a>
```cs
public class EmployeeWithMinMaxWidth
{
    [Column(MinWidth = 40)]
    public required string Name { get; init; }

    [Column(MaxWidth = 20)]
    public required string Email { get; init; }
}
```
<sup><a href='/src/Excelsior.Tests/ColumnWidths.cs#L299-L310' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnMinMaxWidthModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Max Row Height

Cells with wrapped or multi-line content can cause rows to grow very tall. A maximum row height (in points) can be set on the `BookBuilder` (applied to every sheet) or per sheet on `AddSheet`. Rows whose estimated content fits within the limit are left to auto-size; rows that would exceed it are capped.

The header row is exempt — it always auto-sizes, regardless of `MaxRowHeight`.


#### Per-sheet

<!-- snippet: MaxRowHeight -->
<a id='snippet-MaxRowHeight'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(notes, maxRowHeight: 60);
```
<sup><a href='/src/Excelsior.Tests/RowHeights.cs#L34-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-MaxRowHeight' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Book-wide

<!-- snippet: BookMaxRowHeight -->
<a id='snippet-BookMaxRowHeight'></a>
```cs
var builder = new BookBuilder(maxRowHeight: 60);
builder.AddSheet(notes);
```
<sup><a href='/src/Excelsior.Tests/RowHeights.cs#L51-L56' title='Snippet source file'>snippet source</a> | <a href='#snippet-BookMaxRowHeight' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Complex Types

For complex types, by default is to render via `.ToString()`.


#### Models

<!-- snippet: ComplexTypeModels -->
<a id='snippet-ComplexTypeModels'></a>
```cs
public record Person(string Name, Address Address);

public record Address(int Number, string Street, string City, State State, ushort PostCode);
```
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithToString.cs#L10-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeModels' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithToString.cs#L21-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithToString' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/ComplexTypeWithToString.Test_Sheet1.png">


### Custom render for Complex Types

<!-- snippet: ComplexTypeWithCustomRenderInit -->
<a id='snippet-ComplexTypeWithCustomRenderInit'></a>
```cs
[ModuleInitializer]
public static void Init() =>
    ValueRenderer.For<Address>(_ => $"{_.Number}, {_.Street}, {_.City}, {_.State}, {_.PostCode}");
```
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithCustomRender.cs#L14-L20' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRenderInit' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithCustomRender.cs#L25-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithCustomRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result:

<img src="/src/Excelsior.Tests/ComplexTypeWithCustomRender.Test_Sheet1.png">


### Links

Excelsior has first-class support for hyperlinks via the `Link` type.

```cs
namespace Excelsior;

public record Link(string Url, string? Text = null);
```

A single `Link` property renders as a clickable hyperlink. When `Text` is provided it is shown as the display text; otherwise the URL is displayed.

For `IEnumerable<Link>`, items are rendered as blue-styled rich text with URLs visible (e.g., `● Google (https://google.com)`) but are not individually clickable since Excel supports only one hyperlink per cell.


#### Model

<!-- snippet: LinkModel -->
<a id='snippet-LinkModel'></a>
```cs
public record LinkTarget(
    string Name,
    Link Link,
    Link? NullableLink,
    IEnumerable<Link> Links,
    IEnumerable<Link>? NullableLinks,
    IEnumerable<Link?> LinksWithNulls);
```
<sup><a href='/src/Excelsior.Tests/LinkTests.cs#L4-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-LinkModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Build

<!-- snippet: LinkUsage -->
<a id='snippet-LinkUsage'></a>
```cs
List<LinkTarget> data =
[
    new(
        "Test",
        new("https://google.com", "Google"),
        new("https://github.com", "GitHub"),
        [
            new("https://google.com", "Google"),
            new("https://github.com", "GitHub")
        ],
        [
            new("https://google.com", "Google")
        ],
        [
            new("https://google.com", "Google"),
            null,
            new("https://github.com", "GitHub")
        ])
];

var builder = new BookBuilder();
builder.AddSheet(data);
```
<sup><a href='/src/Excelsior.Tests/LinkTests.cs#L19-L44' title='Snippet source file'>snippet source</a> | <a href='#snippet-LinkUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/LinkTests.Test_Sheet1.png">


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

using var book = await builder.Build();
```
<sup><a href='/src/Excelsior.Tests/WhitespaceTests.cs#L7-L24' title='Snippet source file'>snippet source</a> | <a href='#snippet-Whitespace' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/WhitespaceTests.Whitespace_Sheet1.png">


#### Disable whitespace trimming

<!-- snippet: DisableWhitespaceTrim -->
<a id='snippet-DisableWhitespaceTrim'></a>
```cs
static void DisableTrimWhitespace() =>
    ValueRenderer.DisableWhitespaceTrimming();
```
<sup><a href='/src/StaticSettingsTests/DisableWhitespaceTrimmingTests.cs#L12-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-DisableWhitespaceTrim' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/DisableWhitespaceTrimmingTests.Whitespace_Sheet1.png">


### Enumerable string properties

Properties that are castable to an `IEnumerable<string>` will automatically be rendered as a point form list.


#### Model

<!-- snippet: EnumerableModel -->
<a id='snippet-EnumerableModel'></a>
```cs
public record Person(string Name, IEnumerable<string> PhoneNumbers);
```
<sup><a href='/src/Excelsior.Tests/EnumerableStringTests.cs#L4-L8' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableModel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/EnumerableStringTests.cs#L13-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-EnumerableUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/EnumerableStringTests.Test_Sheet1.png">


## Binding Model

The recommended approach is to use a specific type for binding.

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
<sup><a href='/src/Excelsior.Tests/BindingModel.cs#L3-L26' title='Snippet source file'>snippet source</a> | <a href='#snippet-DataModel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/BindingModel.cs#L28-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-EmployeeBindingModel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/BindingModel.cs#L50-L65' title='Snippet source file'>snippet source</a> | <a href='#snippet-ModelProjection' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ColumnAttribute

`ColumnAttribute` allows customization of rendering at the model level.

It is intended as the preferred approach over usage of `DisplayAttribute` and `DisplayNameAttribute`.

`DisplayAttribute` and `DisplayNameAttribute` are support for scenarios where it is not convenient to reference Excelsior from that assembly.<!-- singleLineInclude: DisplayAttributeScenario. path: /docs/DisplayAttributeScenario.include.md -->


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
    public int Width { get; set; } = -1;
    public int MinWidth { get; set; } = -1;
    public int MaxWidth { get; set; } = -1;
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public bool IsHtml { get; set; }

    public bool Filter
    {
        get;
        set
        {
            field = value;
            FilterHasValue = true;
        }
    }

    internal bool FilterHasValue { get; private set; }

    public bool Include
    {
        get;
        set
        {
            field = value;
            IncludeHasValue = true;
        }
    } = true;

    internal bool IncludeHasValue { get; private set; }
}
```
<sup><a href='/src/Excelsior/Attributes/ColumnAttribute.cs#L1-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttribute.cs' title='Start of snippet'>anchor</a></sup>
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

    [Column(Order = 3, NullDisplay = "unknown")]
    public Date? HireDate { get; init; }
}
```
<sup><a href='/src/Excelsior.Tests/ColumnAttributeTests.cs#L4-L21' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttributeModel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Excelsior.Tests/ColumnAttributeTests.cs#L26-L50' title='Snippet source file'>snippet source</a> | <a href='#snippet-ColumnAttribute' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/ColumnAttributeTests.Test_Sheet1.png">


### ValueRenderer.ForEnums

`ValueRenderer.ForEnums` can be used to control the rendering for all enums


#### Config in a ModuleInitializer

<!-- snippet: ValueRendererForEnumsInit -->
<a id='snippet-ValueRendererForEnumsInit'></a>
```cs
static void CustomEnumRender() =>
    ValueRenderer.ForEnums(_ => _.ToString().ToUpper());
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForEnums.cs#L12-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForEnumsInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example use

<!-- snippet: ValueRendererForEnums -->
<a id='snippet-ValueRendererForEnums'></a>
```cs
var builder = new BookBuilder();

List<Car> data =
[
    new()
    {
        Manufacturer = Manufacturer.BuildYourDream,
        Color = Color.AntiqueWhite,
        NullableColor = Color.AntiqueWhite,
    },
    new()
    {
        Manufacturer = Manufacturer.BuildYourDream,
        Color = Color.AntiqueWhite,
    }
];
builder.AddSheet(data);
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForEnums.cs#L22-L42' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForEnums' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/ValueRendererForEnums.Test_Sheet1.png">


### ValueRenderer.ForEnums using Humanizer

Using [Humanizer](https://github.com/Humanizr/Humanizer) to convert enums to strings:


#### Config in a ModuleInitializer

<!-- snippet: ValueRendererForEnumsHumanizerInit -->
<a id='snippet-ValueRendererForEnumsHumanizerInit'></a>
```cs
static void CustomEnumRender() =>
    ValueRenderer.ForEnums(_ => _.Humanize());
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForEnumsHumanizer.cs#L12-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForEnumsHumanizerInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example use

<!-- snippet: ValueRendererForEnumsHumanizer -->
<a id='snippet-ValueRendererForEnumsHumanizer'></a>
```cs
var builder = new BookBuilder();

List<Car> data =
[
    new()
    {
        Manufacturer = Manufacturer.BuildYourDream,
        Color = Color.AntiqueWhite,
        NullableColor = Color.AntiqueWhite,
    },
    new()
    {
        Manufacturer = Manufacturer.BuildYourDream,
        Color = Color.AntiqueWhite,
    }
];
builder.AddSheet(data);
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForEnumsHumanizer.cs#L22-L42' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForEnumsHumanizer' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/ValueRendererForEnumsHumanizer.Test_Sheet1.png">


### ValueRenderer.For&lt;T&gt;

`ValueRenderer.For<T>` can be used to control the rendering for all instances of a specific type. For example rendering `bool` as "Yes"/"No":


#### Config in a ModuleInitializer

<!-- snippet: ValueRendererForBoolInit -->
<a id='snippet-ValueRendererForBoolInit'></a>
```cs
static void CustomBoolRender() =>
    ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForBool.cs#L12-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForBoolInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example use

<!-- snippet: ValueRendererForBool -->
<a id='snippet-ValueRendererForBool'></a>
```cs
var builder = new BookBuilder();

List<Target> data =
[
    new()
    {
        Name = "Alice",
        IsActive = true,
    },
    new()
    {
        Name = "Bob",
        IsActive = false,
    }
];
builder.AddSheet(data);
```
<sup><a href='/src/StaticSettingsTests/ValueRendererForBool.cs#L22-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererForBool' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/ValueRendererForBool.Test_Sheet1.png">


#### Type specificity

When multiple `For<T>` registrations match a property type, the most specific type wins. For example, `For<Color>(...)` takes precedence over `For<Enum>(...)` for `Color` properties, while other enum types still use the `Enum` fallback.


### ValueRenderer.NullDisplayFor&lt;T&gt;

`ValueRenderer.NullDisplayFor<T>` can be used to control the display text when a nullable property is null. This combines well with `ValueRenderer.For<T>`:


#### Config in a ModuleInitializer

<!-- snippet: ValueRendererNullDisplayForBoolInit -->
<a id='snippet-ValueRendererNullDisplayForBoolInit'></a>
```cs
static void CustomBoolRender()
{
    ValueRenderer.For<bool>(_ => _ ? "Yes" : "No");
    ValueRenderer.NullDisplayFor<bool>("Unknown");
}
```
<sup><a href='/src/StaticSettingsTests/ValueRendererNullDisplayForBool.cs#L15-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererNullDisplayForBoolInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example use

<!-- snippet: ValueRendererNullDisplayForBool -->
<a id='snippet-ValueRendererNullDisplayForBool'></a>
```cs
var builder = new BookBuilder();

List<Target> data =
[
    new()
    {
        Name = "Alice",
        IsActive = true,
    },
    new()
    {
        Name = "Bob",
        IsActive = false,
    },
    new()
    {
        Name = "Charlie",
        IsActive = null,
    }
];
builder.AddSheet(data);
```
<sup><a href='/src/StaticSettingsTests/ValueRendererNullDisplayForBool.cs#L28-L52' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererNullDisplayForBool' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/ValueRendererNullDisplayForBool.Test_Sheet1.png">


### ValueRenderer.NullDisplayFor&lt;Enum&gt;

`ValueRenderer.NullDisplayFor<Enum>` can be used to set a default display text for all null enum properties:


#### Config in a ModuleInitializer

<!-- snippet: ValueRendererNullDisplayForEnumInit -->
<a id='snippet-ValueRendererNullDisplayForEnumInit'></a>
```cs
static void CustomNullEnumDisplay() =>
    ValueRenderer.NullDisplayFor<Enum>("Unknown");
```
<sup><a href='/src/StaticSettingsTests/ValueRendererNullDisplayForEnum.cs#L12-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererNullDisplayForEnumInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example use

<!-- snippet: ValueRendererNullDisplayForEnum -->
<a id='snippet-ValueRendererNullDisplayForEnum'></a>
```cs
var builder = new BookBuilder();

List<Target> data =
[
    new()
    {
        Name = "Alice",
        Color = Color.AntiqueWhite,
    },
    new()
    {
        Name = "Bob",
    }
];
builder.AddSheet(data);
```
<sup><a href='/src/StaticSettingsTests/ValueRendererNullDisplayForEnum.cs#L22-L40' title='Snippet source file'>snippet source</a> | <a href='#snippet-ValueRendererNullDisplayForEnum' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/ValueRendererNullDisplayForEnum.Test_Sheet1.png">

The same type specificity applies to `NullDisplayFor<T>`: `NullDisplayFor<Color>("Color unknown")` takes precedence over `NullDisplayFor<Enum>("Enum unknown")` for `Color?` properties.


### Date formats

`DateTime` and `DateOnly` are passed directly in to the respective library.

Excel is directed (using a format string) to render the value using the following:

 * `yyyy-MM-dd HH:mm:ss` for `DateTime`s
 * `yyyy-MM-dd` for `DateOnly`s

Excel has no direct support for `DateTimeOffset`. So `DateTimeOffset`s are stored as strings using the `yyyy-MM-dd HH:mm:ss z` format and `CultureInfo.InvariantCulture`


### Custom Date formats

Date formats can be customized:

<!-- snippet: DateFormatsInit -->
<a id='snippet-DateFormatsInit'></a>
```cs
static void CustomDateFormats()
{
    ValueRenderer.DefaultDateFormat = "yyyy/MM/dd" ;
    ValueRenderer.DefaultDateTimeFormat = "yyyy/MM/dd HH:mm:ss" ;
    ValueRenderer.DefaultDateTimeOffsetFormat = "yyyy/MM/dd HH:mm:ss z" ;
}
```
<sup><a href='/src/StaticSettingsTests/DateFormats.cs#L17-L26' title='Snippet source file'>snippet source</a> | <a href='#snippet-DateFormatsInit' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/StaticSettingsTests/DateFormats.Test_Sheet1.png">


### Filters

By default, auto-filter is enabled on all columns.


#### Disable all filters

<!-- snippet: FilterAllOff -->
<a id='snippet-FilterAllOff'></a>
```cs
var builder = new BookBuilder();
var sheet = builder.AddSheet(Data());
sheet.DisableFilter();
```
<sup><a href='/src/Excelsior.Tests/FilterTests.cs#L32-L38' title='Snippet source file'>snippet source</a> | <a href='#snippet-FilterAllOff' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/FilterTests.AllOff_Sheet1.png">


#### Enable filter on specific columns

Filters can be disabled at the sheet level, then selectively enabled on specific columns:

<!-- snippet: FilterDefaultOffWithOneOn -->
<a id='snippet-FilterDefaultOffWithOneOn'></a>
```cs
var builder = new BookBuilder();
var sheet = builder.AddSheet(Data());
sheet.DisableFilter();
sheet.Filter(_ => _.Name);
```
<sup><a href='/src/Excelsior.Tests/FilterTests.cs#L47-L54' title='Snippet source file'>snippet source</a> | <a href='#snippet-FilterDefaultOffWithOneOn' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/FilterTests.DefaultOffWithOneOn_Sheet1.png">


#### Disable filter on specific columns

Individual columns can opt out of filtering while the rest remain enabled:

<!-- snippet: FilterDefaultOnWithOneOff -->
<a id='snippet-FilterDefaultOnWithOneOff'></a>
```cs
var builder = new BookBuilder();
var sheet = builder.AddSheet(Data());
sheet.Column(
    _ => _.Age,
    _ => _.Filter = false);
```
<sup><a href='/src/Excelsior.Tests/FilterTests.cs#L63-L71' title='Snippet source file'>snippet source</a> | <a href='#snippet-FilterDefaultOnWithOneOff' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/FilterTests.DefaultOnWithOneOff_Sheet1.png">


#### ColumnAttribute

```
public class Employee
{
    [Column(Filter = true)]
    public required string Name { get; init; }
```


### Include/Exclude Columns

Columns can be included or excluded from the output at runtime. This is useful when generating multiple spreadsheets from the same model with different columns based on some state.


#### Toggle based on state

<!-- snippet: IncludeToggleBasedOnState -->
<a id='snippet-IncludeToggleBasedOnState'></a>
```cs
var data = Data();
var isInternalReport = true;

var builder = new BookBuilder();
var sheet = builder.AddSheet(data);
sheet.Include(_ => _.Email, !isInternalReport);
```
<sup><a href='/src/Excelsior.Tests/IncludeTests.cs#L83-L92' title='Snippet source file'>snippet source</a> | <a href='#snippet-IncludeToggleBasedOnState' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/IncludeTests.ToggleBasedOnState_Sheet1.png">


#### Multiple spreadsheets from the same model

The same data can produce different reports by toggling column inclusion per spreadsheet:

<!-- snippet: IncludeMultipleSpreadsheets_Public -->
<a id='snippet-IncludeMultipleSpreadsheets_Public'></a>
```cs
var data = Data();

// Public report: exclude age and email
var builder = new BookBuilder();
var sheet = builder.AddSheet(data);
sheet.Exclude(_ => _.Age);
sheet.Exclude(_ => _.Email);
```
<sup><a href='/src/Excelsior.Tests/IncludeTests.cs#L101-L111' title='Snippet source file'>snippet source</a> | <a href='#snippet-IncludeMultipleSpreadsheets_Public' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/IncludeTests.MultipleSpreadsheetsSameModel_Public_Sheet1.png">

<!-- snippet: IncludeMultipleSpreadsheets_Internal -->
<a id='snippet-IncludeMultipleSpreadsheets_Internal'></a>
```cs
List<Target> data = [
    new("Alice", 30, "alice@test.com"),
    new("Bob", 25, "bob@test.com")
];

// Internal report: include all columns
var builder = new BookBuilder();
builder.AddSheet(data);
```
<sup><a href='/src/Excelsior.Tests/IncludeTests.cs#L120-L131' title='Snippet source file'>snippet source</a> | <a href='#snippet-IncludeMultipleSpreadsheets_Internal' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/IncludeTests.MultipleSpreadsheetsSameModel_Internal_Sheet1.png">


#### Exclude a column

<!-- snippet: IncludeExcludeOne -->
<a id='snippet-IncludeExcludeOne'></a>
```cs
List<Target> data = [
    new("Alice", 30, "alice@test.com"),
    new("Bob", 25, "bob@test.com")
];
var builder = new BookBuilder();
var sheet = builder.AddSheet(data);
sheet.Exclude(_ => _.Age);
```
<sup><a href='/src/Excelsior.Tests/IncludeTests.cs#L29-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-IncludeExcludeOne' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/IncludeTests.ExcludeOne_Sheet1.png">


#### Exclude via Column configuration

<!-- snippet: IncludeExcludeOneViaColumn -->
<a id='snippet-IncludeExcludeOneViaColumn'></a>
```cs
List<Target> data = [
    new("Alice", 30, "alice@test.com"),
    new("Bob", 25, "bob@test.com")
];
var builder = new BookBuilder();
var sheet = builder.AddSheet(data);
sheet.Column(
    _ => _.Age,
    _ => _.Include = false);
```
<sup><a href='/src/Excelsior.Tests/IncludeTests.cs#L48-L60' title='Snippet source file'>snippet source</a> | <a href='#snippet-IncludeExcludeOneViaColumn' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/IncludeTests.ExcludeOneViaColumn_Sheet1.png">


#### ColumnAttribute

```
public class Employee
{
    [Column(Include = false)]
    public required string Name { get; init; }
```


### Splitting


`SplitAttribute` can be used push properties up.

<!-- snippet: ComplexTypeWithSplitter -->
<a id='snippet-ComplexTypeWithSplitter'></a>
```cs
public record Person(
    string Name,
    [Split] Address Address);

public record Address(int StreetNumber, string Street, string City, State State, ushort PostCode);
```
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithSplitter.cs#L10-L18' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithSplitter' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/ComplexTypeWithSplitter.Test_Sheet1.png">


#### UseHierachyForName

`SplitAttribute.UseHierachyForName` can be used to prefix members with the parent property name.

<!-- snippet: ComplexTypeWithSplitterUseHierachyForName -->
<a id='snippet-ComplexTypeWithSplitterUseHierachyForName'></a>
```cs
public record Person(
    string Name,
    [Split(UseHierachyForName = true)]
    Address Address);

public record Address(int Number, string Street, string City, State State, ushort PostCode);
```
<sup><a href='/src/Excelsior.Tests/ComplexTypeWithSplitterUseHierachyForName.cs#L10-L19' title='Snippet source file'>snippet source</a> | <a href='#snippet-ComplexTypeWithSplitterUseHierachyForName' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<img src="/src/Excelsior.Tests/ComplexTypeWithSplitterUseHierachyForName.Test_Sheet1.png">


### Source Generated Extensions

A source generator is included that generates typed extension methods for `ISheetBuilder`, providing a more concise API.


#### Model

Add `[SheetModel]` to the model class:

<!-- snippet: SourceGeneratedModel -->
<a id='snippet-SourceGeneratedModel'></a>
```cs
[SheetModel]
public class GeneratedTestModel
{
    public required string Name { get; init; }
    public required int Age { get; init; }
}
```
<sup><a href='/src/Excelsior.Tests/SourceGeneratorIntegrationTests.cs#L60-L69' title='Snippet source file'>snippet source</a> | <a href='#snippet-SourceGeneratedModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

This generates typed extension methods for each property, such as `NameColumn`, `NameOrder`, `AgeWidth`, etc.


#### Usage

Instead of:

```cs
sheet.Column(_ => _.Name, _ => _.Heading = "Full Name");
sheet.Order(_ => _.Age, 1);
```

Use the generated methods:

<!-- snippet: SourceGeneratedUsage -->
<a id='snippet-SourceGeneratedUsage'></a>
```cs
var builder = new BookBuilder();

List<GeneratedTestModel> data =
[
    new() { Name = "Alice", Age = 30 },
    new() { Name = "Bob", Age = 25 },
];

var sheet = builder.AddSheet(data);
sheet.NameColumn(_ => _.Heading = "Full Name");
sheet.AgeOrder(1);
sheet.NameOrder(2);
sheet.AgeWidth(15);
```
<sup><a href='/src/Excelsior.Tests/SourceGeneratorIntegrationTests.cs#L7-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-SourceGeneratedUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Generated methods per property

For each public property, the following extension methods are generated:

 * `{Property}Column` — configure the column (heading, order, width, style, etc.)
 * `{Property}HeadingText` — set the heading text
 * `{Property}Order` — set the column order
 * `{Property}Width` — set the column width
 * `{Property}HeadingStyle` — set the heading style
 * `{Property}CellStyle` — set the cell style
 * `{Property}Format` — set the format string
 * `{Property}NullDisplay` — set the null display text
 * `{Property}IsHtml` — mark the column as HTML
 * `{Property}Render` — set a custom render function
 * `{Property}Filter` — enable auto-filter for the column
 * `{Property}Include` — include or exclude the column from the output
 * `{Property}Exclude` — exclude the column from the output

Properties with `[Ignore]` are skipped. Properties with `[Split]` (or types with `[Split]`) are recursed into, generating methods for the nested properties.


## Word Tables

`WordTableBuilder<TModel>` renders model data into a Word `<w:tbl>` element that can be appended to an existing Word document. It reuses the same property discovery, column ordering, and per-column configuration as `BookBuilder`.

<!-- snippet: WordTableUsage -->
<a id='snippet-WordTableUsage'></a>
```cs
var builder = new WordTableBuilder<Employee>(employees);

using var stream = new MemoryStream();
using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
{
    var mainPart = doc.AddMainDocumentPart();
    mainPart.Document = new(new Body());

    var table = builder.Build(mainPart);
    var body = mainPart.Document.Body!;
    body.Append(table);
```
<sup><a href='/src/Excelsior.Tests/Word/WordTableBuilderTests.cs#L13-L27' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordTableUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Column configuration (headings, ordering, render, etc.) works the same as with `BookBuilder`:

```cs
var builder = new WordTableBuilder<Employee>(employees)
    .Column(
        _ => _.Name,
        _ => _.Heading = "Person");
```

When a `MainDocumentPart` is passed to `Build()`, `Link`-typed properties produce real `<w:hyperlink>` elements. When omitted, links fall back to their display text.


### Limitations

Formula columns are not supported in Word tables. Word has no equivalent of Excel cell formulas, so configuring a `Formula` on a column used with `WordTableBuilder` will throw when `Build()` is called. Use `Render` or a computed property instead.


## Icon

[Grim Fandango](https://github.com/PapirusDevelopmentTeam/papirus-icon-theme/blob/master/Papirus/64x64/apps/grim-fandango-remastered.svg) from [Papirus Icons](https://github.com/PapirusDevelopmentTeam/papirus-icon-theme).

The [Excelsior Line](https://grim-fandango.fandom.com/wiki/Excelsior_Line) is a travel package sold by Manuel Calavera in the Lucas Arts game "Grim Fandango". The package consists of nothing more than a walking stick with a compass in the handle.
