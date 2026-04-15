using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

[TestFixture]
public class WordTableBuilderTests
{
    [Test]
    public async Task RendersTableFromAttributedModel()
    {
        var employees = SampleData.Employees();
        var builder = new WordTableBuilder<Employee>(employees);

        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());

            var table = builder.Build(mainPart);
            mainPart.Document.Body!.Append(table);
            mainPart.Document.Body.Append(new SectionProperties(
                new PageSize
                {
                    Width = 12240,
                    Height = 15840
                },
                new PageMargin
                {
                    Top = 1440,
                    Right = 1440,
                    Bottom = 1440,
                    Left = 1440,
                    Header = 720,
                    Footer = 720
                }));
        }

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public void HeaderRowIsBoldAndCentered()
    {
        var builder = new WordTableBuilder<Employee>(SampleData.Employees());
        var table = builder.Build();

        var headerRow = table.Elements<TableRow>().First();
        var headerParagraph = headerRow.Elements<TableCell>().First().GetFirstChild<Paragraph>()!;
        var justification = headerParagraph.ParagraphProperties!.GetFirstChild<Justification>()!;
        AreEqual(JustificationValues.Center, justification.Val?.Value);

        var headerRun = headerParagraph.GetFirstChild<Run>()!;
        IsNotNull(headerRun.RunProperties!.GetFirstChild<Bold>());
    }

    [Test]
    public void DataRowsMatchEmployeeCount()
    {
        var employees = SampleData.Employees();
        var table = new WordTableBuilder<Employee>(employees).Build();

        var rows = table.Elements<TableRow>().ToList();
        AreEqual(employees.Count + 1, rows.Count); // +1 for header row
    }

    [Test]
    public void ColumnHeadingsHonorColumnAttribute()
    {
        var table = new WordTableBuilder<Employee>([]).Build();
        var headerCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
        var headings = headerCells
            .Select(_ => _.GetFirstChild<Paragraph>()!.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text)
            .ToList();

        // Employee model declares Order=1..5 with explicit Headings; IsActive/Status fall after.
        AreEqual("Employee ID", headings[0]);
        AreEqual("Full Name", headings[1]);
        AreEqual("Email Address", headings[2]);
    }

    public class LinkRow
    {
        public required string Label { get; init; }
        public required Link Site { get; init; }
    }

    [Test]
    public void LinkValueProducesHyperlinkWhenMainPartGiven()
    {
        var rows = new[]
        {
            new LinkRow
            {
                Label = "Anthropic",
                Site = new("https://www.anthropic.com", "Home")
            }
        };

        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new(new Body());

        var table = new WordTableBuilder<LinkRow>(rows).Build(mainPart);

        var cells = table.Elements<TableRow>().Skip(1).First().Elements<TableCell>().ToList();
        var linkCell = cells[1];
        var hyperlink = linkCell.GetFirstChild<Paragraph>()!.GetFirstChild<Hyperlink>();
        IsNotNull(hyperlink);

        var rel = mainPart.HyperlinkRelationships.Single();
        AreEqual("https://www.anthropic.com/", rel.Uri.ToString());
        AreEqual(rel.Id, hyperlink!.Id?.Value);

        var run = hyperlink.GetFirstChild<Run>()!;
        AreEqual("Home", run.GetFirstChild<Text>()!.Text);
        IsNotNull(run.RunProperties!.GetFirstChild<Color>());
        IsNotNull(run.RunProperties.GetFirstChild<Underline>());
    }

    [Test]
    public void LinkValueFallsBackToTextWhenMainPartOmitted()
    {
        var rows = new[]
        {
            new LinkRow
            {
                Label = "Anthropic",
                Site = new("https://www.anthropic.com", "Home")
            }
        };

        var table = new WordTableBuilder<LinkRow>(rows).Build();

        var cells = table.Elements<TableRow>().Skip(1).First().Elements<TableCell>().ToList();
        var linkCell = cells[1];
        var paragraph = linkCell.GetFirstChild<Paragraph>()!;
        IsNull(paragraph.GetFirstChild<Hyperlink>());

        var run = paragraph.GetFirstChild<Run>()!;
        AreEqual("Home", run.GetFirstChild<Text>()!.Text);
    }

    [Test]
    public void FluentColumnConfigurationOverridesHeading()
    {
        var builder = new WordTableBuilder<Employee>([])
            .Column(
                _ => _.Name,
                _ => _.Heading = "Person");

        var table = builder.Build();
        var headerCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
        var headings = headerCells
            .Select(_ => _.GetFirstChild<Paragraph>()!.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text)
            .ToList();

        IsTrue(headings.Contains("Person"));
        IsFalse(headings.Contains("Full Name"));
    }
}
