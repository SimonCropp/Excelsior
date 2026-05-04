[TestFixture]
public class BookReaderDelegateTests
{
    public class Source
    {
        public required string Code { get; init; }
        public required string Priority { get; init; }
    }

    public class Target
    {
        public string Code { get; set; } = "";
        public Priority Priority { get; set; }
    }

    public enum Priority
    {
        Low,
        Medium,
        High
    }

    [Test]
    public async Task DelegateConversion_StrongTyped()
    {
        var source = new[]
        {
            new Source { Code = "A", Priority = "low" },
            new Source { Code = "B", Priority = "HIGH" },
            new Source { Code = "C", Priority = "medium" }
        };

        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(source);
        await builder.ToStream(stream);
        stream.Position = 0;

        #region ReaderDelegate

        var reader = new BookReader();
        var sheet = reader.AddSheet<Target>();
        sheet.Convert(
            _ => _.Priority,
            cell =>
            {
                var raw = cell.InnerText.Trim().ToLowerInvariant();
                return raw switch
                {
                    "low" => Priority.Low,
                    "medium" => Priority.Medium,
                    "high" => Priority.High,
                    _ => Priority.Low
                };
            });
        reader.Convert(stream);

        #endregion

        Assert.That(sheet.Rows.Select(_ => _.Priority), Is.EqualTo([Priority.Low, Priority.High, Priority.Medium]));
        Assert.That(sheet.Rows.Select(_ => _.Code), Is.EqualTo(["A", "B", "C"]));
    }

    [Test]
    public async Task DelegateConversion_Dictionary()
    {
        var source = new[]
        {
            new Source { Code = "A", Priority = "low" },
            new Source { Code = "B", Priority = "high" }
        };

        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(source);
        await builder.ToStream(stream);
        stream.Position = 0;

        #region ReaderDictionaryDelegate

        var reader = new BookReader();
        var sheet = reader.AddSheet();
        sheet.Column<string>("Code");
        sheet.Column(
            "Priority",
            cell =>
            {
                var text = cell.InnerText;
                return text.Trim().ToLowerInvariant() switch
                {
                    "low" => 1,
                    "medium" => 2,
                    "high" => 3,
                    _ => 0
                };
            });
        reader.Convert(stream);

        #endregion

        Assert.That(sheet.Rows[0]["Code"], Is.EqualTo("A"));
        Assert.That(sheet.Rows[0]["Priority"], Is.EqualTo(1));
        Assert.That(sheet.Rows[1]["Code"], Is.EqualTo("B"));
        Assert.That(sheet.Rows[1]["Priority"], Is.EqualTo(3));
    }
}
