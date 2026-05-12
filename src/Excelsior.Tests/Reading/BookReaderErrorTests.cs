[TestFixture]
public class BookReaderErrorTests
{
    public class StringSource
    {
        public required string Number { get; init; }
    }

    public class IntTarget
    {
        public int Number { get; set; }
    }

    static async Task<MemoryStream> WriteStringNumber()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet<StringSource>(
        [
            new() { Number = "42" },
            new() { Number = "not-a-number" },
            new() { Number = "100" }
        ]);
        await builder.ToStream(stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task Convert_Throws_ReadException_With_Errors()
    {
        var stream = await WriteStringNumber();

        var reader = new BookReader();
        reader.AddSheet<IntTarget>();

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        var errors = exception.Errors;
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].ColumnName, Is.EqualTo("Number"));
        Assert.That(errors[0].Message, Does.Contain("not-a-number"));
    }

    [Test]
    public async Task TryConvert_Returns_Errors_Without_Throwing()
    {
        #region BookReaderTryConvert

        var stream = await WriteStringNumber();
        var reader = new BookReader();
        var sheet = reader.AddSheet<IntTarget>();

        var result = reader.TryConvert(stream);
        if (!result)
        {
            foreach (var error in result.Errors)
            {
                Console.WriteLine(error);
            }
        }

        #endregion

        Assert.That(result.Succeeded, Is.False);
        Assert.That((bool)result, Is.False);
        var errors = (ReadError[])result;
        Assert.That(errors, Has.Length.EqualTo(1));
        Assert.That(errors[0].ColumnName, Is.EqualTo("Number"));
        Assert.That(sheet.Rows, Has.Count.EqualTo(3));
        Assert.That(sheet.Rows[0].Number, Is.EqualTo(42));
        Assert.That(sheet.Rows[1].Number, Is.EqualTo(0));
        Assert.That(sheet.Rows[2].Number, Is.EqualTo(100));
    }

    [Test]
    public async Task Exception_Errors_Match_TryConvert_Errors()
    {
        var stream1 = await WriteStringNumber();
        var stream2 = await WriteStringNumber();

        var reader1 = new BookReader();
        reader1.AddSheet<IntTarget>();
        var tryResult = reader1.TryConvert(stream1);

        var reader2 = new BookReader();
        reader2.AddSheet<IntTarget>();
        var exception = Assert.Throws<ReadException>(() => reader2.Convert(stream2))!;

        Assert.That(exception.Errors.Select(_ => _.Message), Is.EqualTo(tryResult.Errors.Select(_ => _.Message)));
    }

    public class OneCol
    {
        public required string A { get; init; }
    }

    public class TwoCols
    {
        public required string A { get; init; }
        public required string B { get; init; }
    }

    public class ThreeCols
    {
        public required string A { get; init; }
        public required string B { get; init; }
        public required string C { get; init; }
    }

    static async Task<MemoryStream> Write<T>(params IEnumerable<T> rows)
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(rows);
        await builder.ToStream(stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task ColumnMismatch_StrongTyped_OneMissing()
    {
        var stream = await Write(new OneCol { A = "x" });

        var reader = new BookReader();
        reader.AddSheet<TwoCols>();

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        Assert.That(exception.Errors, Has.Count.EqualTo(1));
        Assert.That(exception.Errors[0].ColumnName, Is.EqualTo("B"));
        Assert.That(exception.Errors[0].Message, Does.Contain("not found"));
    }

    [Test]
    public async Task ColumnMismatch_StrongTyped_MultipleMissingProduceMultipleErrors()
    {
        var stream = await Write(new OneCol { A = "x" });

        var reader = new BookReader();
        reader.AddSheet<ThreeCols>();

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        Assert.That(exception.Errors, Has.Count.EqualTo(2));
        Assert.That(exception.Errors.Select(_ => _.ColumnName), Is.EquivalentTo(["B", "C"]));
    }

    [Test]
    public async Task ColumnMismatch_Dictionary_HeadingNotInFile()
    {
        var stream = await Write(new OneCol { A = "x" });

        var reader = new BookReader();
        var sheet = reader.AddSheet();
        sheet.Column<string>("Nope");

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        Assert.That(exception.Errors, Has.Count.EqualTo(2));
        Assert.That(exception.Errors.Any(_ => _.ColumnName == "Nope" && _.Message.Contains("not found")));
        Assert.That(exception.Errors.Any(_ => _.Message.Contains("Unrecognized header 'A'")));
    }

    [Test]
    public async Task ColumnMismatch_StopsBeforeRowParsing_NoPerCellErrors()
    {
        // OneCol writes string "A"; reader expects an int "Value" column that
        // doesn't exist. Should produce one missing-column error plus one
        // unrecognized-header error and skip row parsing entirely.
        var stream = await Write<OneCol>(
            new() { A = "x" },
            new() { A = "y" },
            new() { A = "z" });

        var reader = new BookReader();
        var sheet = reader.AddSheet();
        sheet.Column<int>("Value");

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        Assert.That(exception.Errors, Has.Count.EqualTo(2));
        Assert.That(exception.Errors.Any(_ => _.ColumnName == "Value" && _.Message.Contains("not found")));
        Assert.That(exception.Errors.Any(_ => _.Message.Contains("Unrecognized header 'A'")));
        Assert.That(sheet.Rows, Is.Empty);
    }

    public class StringRow
    {
        public required string Value { get; init; }
    }

    public class IntRow
    {
        public int Value { get; set; }
    }

    [Test]
    public async Task ColumnMismatch_DoesNotStopSubsequentSheets()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet([new OneCol { A = "x" }], "First");
        builder.AddSheet([new OneCol { A = "y" }], "Second");
        await builder.ToStream(stream);
        stream.Position = 0;

        var reader = new BookReader();
        var first = reader.AddSheet<TwoCols>("First");
        var second = reader.AddSheet<OneCol>("Second");

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        // Mismatch in "First" emits one error and skips its rows;
        // "Second" still parses successfully.
        Assert.That(exception.Errors, Has.Count.EqualTo(1));
        Assert.That(exception.Errors[0].SheetName, Is.EqualTo("First"));
        Assert.That(exception.Errors[0].ColumnName, Is.EqualTo("B"));
        Assert.That(first.Rows, Is.Empty);
        Assert.That(second.Rows, Has.Count.EqualTo(1));
        Assert.That(second.Rows[0].A, Is.EqualTo("y"));
    }

    [Test]
    public async Task PerCellErrors_AreNotDeduppedAfterColumnsResolve()
    {
        // Column "Value" matches in both; per-cell parse failures must surface
        // as one error per failing row — no deduplication.
        var stream = await Write<StringRow>(
            new() { Value = "x" },
            new() { Value = "y" },
            new() { Value = "z" });

        var reader = new BookReader();
        reader.AddSheet<IntRow>();

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        Assert.That(exception.Errors, Has.Count.EqualTo(3));
        Assert.That(exception.Errors.Select(_ => _.ColumnName), Is.EqualTo(["Value", "Value", "Value"]));
        Assert.That(exception.Errors.Select(_ => _.RowIndex), Is.EqualTo([2, 3, 4]));
    }
}
