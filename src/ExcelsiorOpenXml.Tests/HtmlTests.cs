[TestFixture]
public class HtmlTests
{
    [Test]
    public async Task BasicHtmlFormatting()
    {
        var data = new[]
        {
            new { Id = 1, Description = "<b>Bold text</b> and <i>italic text</i>" },
            new { Id = 2, Description = "<u>Underlined</u> and <font color=\"red\">red text</font>" },
            new { Id = 3, Description = "<b><i>Bold and italic</i></b> combined" }
        };

        var builder = new BookBuilder();
        var sheetBuilder = builder.AddSheet(data);
        sheetBuilder.Column(_ => _.Description, _ => _.IsHtml = true);

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and verify
        using var wb = new XLWorkbook(tempPath);
        await Verify(wb);

        // Cleanup
        File.Delete(tempPath);
    }

    [Test]
    public async Task HtmlWithLineBreaks()
    {
        var data = new[]
        {
            new { Id = 1, Description = "Line 1<br>Line 2<br>Line 3" },
            new { Id = 2, Description = "<b>Bold line 1</b><br><i>Italic line 2</i>" }
        };

        var builder = new BookBuilder();
        var sheetBuilder = builder.AddSheet(data);
        sheetBuilder.Column(_ => _.Description, _ => _.IsHtml = true);

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and verify
        using var wb = new XLWorkbook(tempPath);
        await Verify(wb);

        // Cleanup
        File.Delete(tempPath);
    }

    [Test]
    public async Task HtmlWithColors()
    {
        var data = new[]
        {
            new { Id = 1, Description = "<font color=\"#FF0000\">Red</font> <font color=\"#00FF00\">Green</font> <font color=\"#0000FF\">Blue</font>" },
            new { Id = 2, Description = "<font color=\"red\">Named red</font> and <font color=\"blue\">Named blue</font>" }
        };

        var builder = new BookBuilder();
        var sheetBuilder = builder.AddSheet(data);
        sheetBuilder.Column(_ => _.Description, _ => _.IsHtml = true);

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and verify
        using var wb = new XLWorkbook(tempPath);
        await Verify(wb);

        // Cleanup
        File.Delete(tempPath);
    }

    [Test]
    public async Task ComplexHtmlFormatting()
    {
        var data = new[]
        {
            new {
                Id = 1,
                Description = "<b>Project Status:</b> <font color=\"green\">On Track</font><br>" +
                              "<i>Last updated:</i> <u>2024-01-15</u><br>" +
                              "<b><font color=\"red\">Critical:</font></b> Review required"
            }
        };

        var builder = new BookBuilder();
        var sheetBuilder = builder.AddSheet(data);
        sheetBuilder.Column(_ => _.Description, _ => _.IsHtml = true);

        var document = await builder.Build();

        // Save to temp file
        var tempPath = Path.GetTempFileName() + ".xlsx";
        await using (var stream = File.Create(tempPath))
        {
            await builder.ToStream(stream);
        }

        // Read with ClosedXML and verify
        using var wb = new XLWorkbook(tempPath);
        await Verify(wb);

        // Cleanup
        File.Delete(tempPath);
    }
}
