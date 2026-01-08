# ExcelsiorOpenXml

Excel export library using the raw OpenXML SDK (DocumentFormat.OpenXml) with **HTML content support**.

## Key Features

- ✅ **HTML Content in Cells** - Unlike ClosedXML, supports rich HTML formatting (`<b>`, `<i>`, `<u>`, `<font color>`, `<br>`)
- ✅ **DOM-based Approach** - Simple in-memory manipulation for straightforward usage
- ✅ **Auto-sizing** - Automatic column width and row height estimation
- ✅ **Date/Number Formatting** - Proper Excel number formats for dates and numbers
- ✅ **Sheet Features** - Frozen headers, auto-filter, global styling
- ✅ **Lightweight** - Direct OpenXML SDK usage without wrapper dependencies

## Installation

```bash
dotnet add package ExcelsiorOpenXml
```

## Basic Usage

<!-- snippet: Usage -->
<a id='snippet-usage'></a>
```cs
var builder = new BookBuilder();
builder.AddSheet(employees);
```
<sup><a href='/src/ExcelsiorOpenXml.Tests/UsageTests.cs#L9-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-usage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Save to file:

```csharp
await builder.ToFile("employees.xlsx");
```

Save to stream:

```csharp
await using var stream = new FileStream("employees.xlsx", FileMode.Create);
await builder.ToStream(stream);
```

## HTML Content Support

This is the **key differentiator** from ClosedXML. ExcelsiorOpenXml supports HTML content in cells:

```csharp
var data = new[]
{
    new {
        Id = 1,
        Status = "<b>Project Status:</b> <font color=\"green\">On Track</font><br>" +
                 "<i>Last updated:</i> <u>2024-01-15</u><br>" +
                 "<b><font color=\"red\">Critical:</font></b> Review required"
    }
};

var builder = new BookBuilder();
var sheetBuilder = builder.AddSheet(data);
sheetBuilder.Column(_ => _.Status, _ => _.IsHtml = true);

await builder.ToFile("status.xlsx");
```

Supported HTML tags:
- `<b>`, `<strong>` - Bold text
- `<i>`, `<em>` - Italic text
- `<u>` - Underlined text
- `<br>` - Line breaks
- `<font color="...">` - Text color (hex colors like `#FF0000` or named colors like `red`, `blue`, `green`)

## Styling

### Global Styling

Apply styles to all cells:

```csharp
var builder = new BookBuilder(
    globalStyle: style =>
    {
        style.Font.Bold = true;
        style.Font.Color = System.Drawing.Color.White;
        style.Fill.BackgroundColor = System.Drawing.Color.DarkBlue;
    });
```

### Alternating Row Colors

```csharp
var builder = new BookBuilder(
    useAlternatingRowColors: true,
    alternateRowColor: System.Drawing.Color.LightGray);
```

### Heading Styles

```csharp
var builder = new BookBuilder(
    headingStyle: style =>
    {
        style.Font.Bold = true;
        style.Fill.BackgroundColor = System.Drawing.Color.Navy;
        style.Font.Color = System.Drawing.Color.White;
    });
```

## Sheet Features

### Frozen Headers

Headers are frozen by default. The first row stays visible when scrolling.

### Auto-Filter

Auto-filter is enabled by default, allowing users to filter data in Excel.

### Column Configuration

```csharp
var sheetBuilder = builder.AddSheet(employees);

// Set custom column width
sheetBuilder.Column(_ => _.Description, config =>
{
    config.Width = 50;
});

// Mark column as HTML
sheetBuilder.Column(_ => _.Notes, config =>
{
    config.IsHtml = true;
});
```

## Date and Number Formatting

Dates and numbers are automatically formatted with Excel number formats:

```csharp
var data = new[]
{
    new { Id = 1, Date = new DateTime(2020, 1, 15), Amount = 75000.50m },
    new { Id = 2, Date = new DateTime(2021, 5, 20), Amount = 120000.75m }
};

var builder = new BookBuilder();
builder.AddSheet(data);
// Dates will display as "2020-01-15 00:00:00" with proper Excel date formatting
// Numbers will display with appropriate decimal formatting
```

## Architecture

ExcelsiorOpenXml uses a **DOM (Document Object Model) approach** with the OpenXML SDK:

1. Creates SpreadsheetDocument in memory
2. Builds worksheet structure with cells, rows, and columns
3. Applies styles via a stylesheet with deduplication
4. Supports random access to cells for flexible manipulation

This differs from a SAX (streaming) approach and provides:
- ✅ Simpler code and easier to understand
- ✅ Full access to all cells for modifications
- ✅ Better for small to medium datasets (< 100k rows)

For very large datasets (> 100k rows), consider ExcelsiorAspose or ExcelsiorSyncfusion which have more optimized large-file handling.

## Comparison with Other Implementations

| Feature | ExcelsiorOpenXml | ExcelsiorClosedXml | ExcelsiorAspose |
|---------|------------------|-------------------|----------------|
| HTML Content | ✅ Full support | ❌ Not supported | ✅ Full support |
| License | Free (MIT) | Free (MIT) | Commercial |
| Dependencies | OpenXML SDK only | ClosedXML wrapper | Aspose.Cells |
| Performance | Good | Good | Excellent |
| File Size | Smallest | Medium | Largest |

## Limitations

- **Width/Height Estimation**: Column widths and row heights are estimated using character counts and may not be pixel-perfect
- **Memory Usage**: DOM approach loads entire document in memory
- **HTML Support**: Supports basic HTML tags only (not full HTML rendering)

## Contributing

This library is part of the [Excelsior](https://github.com/SimonCropp/Excelsior) project. See the main repository for contribution guidelines.

## License

MIT License - see LICENSE file for details
