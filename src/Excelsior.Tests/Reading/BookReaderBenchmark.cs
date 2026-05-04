using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Manual perf baseline. Run via:
//   dotnet test src/Excelsior.Tests --filter "FullyQualifiedName~BookReaderBenchmark" -- NUnit.RunInExplicitMode=true
[TestFixture, Explicit("Manual perf benchmark")]
public class BookReaderBenchmark
{
    const int RowCount = 10_000;
    const int Iterations = 11;

    public class ReflectionModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public Date? HireDate { get; set; }
        public int Salary { get; set; }
        public bool IsActive { get; set; }
        public BenchmarkStatus Status { get; set; }
    }

    [SheetModel]
    public class SourceGenModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public Date? HireDate { get; set; }
        public int Salary { get; set; }
        public bool IsActive { get; set; }
        public BenchmarkStatus Status { get; set; }
    }

    public enum BenchmarkStatus
    {
        Active,
        Inactive,
        Pending
    }

    static IEnumerable<ReflectionModel> GenerateReflection()
    {
        for (var i = 0; i < RowCount; i++)
        {
            yield return new()
            {
                Id = i,
                Name = $"Person {i}",
                Email = $"person{i}@example.com",
                HireDate = new Date(2020, 1, 1).AddDays(i % 365),
                Salary = 50_000 + i,
                IsActive = i % 2 == 0,
                Status = (BenchmarkStatus)(i % 3)
            };
        }
    }

    static IEnumerable<SourceGenModel> GenerateSourceGen()
    {
        for (var i = 0; i < RowCount; i++)
        {
            yield return new()
            {
                Id = i,
                Name = $"Person {i}",
                Email = $"person{i}@example.com",
                HireDate = new Date(2020, 1, 1).AddDays(i % 365),
                Salary = 50_000 + i,
                IsActive = i % 2 == 0,
                Status = (BenchmarkStatus)(i % 3)
            };
        }
    }

    static async Task<byte[]> WriteWorkbook<T>(IEnumerable<T> rows)
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(rows);
        await builder.ToStream(stream);
        return stream.ToArray();
    }

    static (TimeSpan min, TimeSpan median, TimeSpan max) MeasureReads(byte[] data, Action<MemoryStream> read)
    {
        // warmup
        using (var warm = new MemoryStream(data, writable: false))
        {
            read(warm);
        }

        var samples = new TimeSpan[Iterations];
        for (var i = 0; i < Iterations; i++)
        {
            using var stream = new MemoryStream(data, writable: false);
            var sw = Stopwatch.StartNew();
            read(stream);
            sw.Stop();
            samples[i] = sw.Elapsed;
        }

        Array.Sort(samples);
        return (samples[0], samples[samples.Length / 2], samples[^1]);
    }

    static void Report(string label, TimeSpan min, TimeSpan median, TimeSpan max)
    {
        var rowsPerSec = RowCount / median.TotalSeconds;
        TestContext.Out.WriteLine(
            $"{label,-30}  min={min.TotalMilliseconds,8:F1} ms  median={median.TotalMilliseconds,8:F1} ms  max={max.TotalMilliseconds,8:F1} ms  rows/s={rowsPerSec,10:F0}");
    }

    [Test]
    public async Task Bench_TypedReflection()
    {
        var data = await WriteWorkbook(GenerateReflection());
        TestContext.Out.WriteLine($"Workbook size: {data.Length:N0} bytes, {RowCount:N0} rows");

        var (min, median, max) = MeasureReads(data, stream =>
        {
            var reader = new BookReader();
            var sheet = reader.AddSheet<ReflectionModel>();
            reader.Convert(stream);
            if (sheet.Rows.Count != RowCount)
            {
                throw new($"Expected {RowCount} rows, got {sheet.Rows.Count}");
            }
        });

        Report("Typed reflection", min, median, max);
    }

    [Test]
    public async Task Bench_TypedSourceGen()
    {
        Assert.That(GeneratedActivators.TryGet<SourceGenModel>(), Is.Not.Null, "Source-gen activator missing for SourceGenModel");
        Assert.That(GeneratedRowReaders.TryGet<SourceGenModel>(), Is.Not.Null, "Source-gen row reader missing for SourceGenModel");

        var data = await WriteWorkbook(GenerateSourceGen());
        TestContext.Out.WriteLine($"Workbook size: {data.Length:N0} bytes, {RowCount:N0} rows");

        var (min, median, max) = MeasureReads(data, stream =>
        {
            var reader = new BookReader();
            var sheet = reader.AddSheet<SourceGenModel>();
            reader.Convert(stream);
            if (sheet.Rows.Count != RowCount)
            {
                throw new($"Expected {RowCount} rows, got {sheet.Rows.Count}");
            }
        });

        Report("Typed source-gen", min, median, max);
    }

    [Test]
    public async Task Bench_OpenXmlOnly()
    {
        // Pure OpenXml DOM cost: open the document, walk every row + cell, read
        // the raw text. No conversion, no model construction. Tells us how
        // much of the total budget is unavoidable plumbing vs. our code.
        var data = await WriteWorkbook(GenerateReflection());
        TestContext.Out.WriteLine($"Workbook size: {data.Length:N0} bytes, {RowCount:N0} rows");

        var (min, median, max) = MeasureReads(data, stream =>
        {
            using var document = SpreadsheetDocument.Open(stream, false);
            var workbookPart = document.WorkbookPart!;
            var sharedStrings = CellConverter.BuildSharedStrings(workbookPart.SharedStringTablePart?.SharedStringTable);
            var sheet = workbookPart.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>().First();
            var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet;
            var sheetData = ws!.GetFirstChild<SheetData>()!;

            var cellCount = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var raw = CellConverter.ReadRaw(cell, sharedStrings);
                    if (raw != null)
                    {
                        cellCount++;
                    }
                }
            }

            if (cellCount < RowCount)
            {
                throw new($"Expected at least {RowCount} cells, got {cellCount}");
            }
        });

        Report("OpenXml iterate+ReadRaw only", min, median, max);
    }

    [Test]
    public async Task Bench_Dictionary()
    {
        var data = await WriteWorkbook(GenerateReflection());
        TestContext.Out.WriteLine($"Workbook size: {data.Length:N0} bytes, {RowCount:N0} rows");

        var (min, median, max) = MeasureReads(data, stream =>
        {
            var reader = new BookReader();
            var sheet = reader.AddSheet();
            sheet
                .Column<int>("Id")
                .Column<string>("Name")
                .Column<string>("Email")
                .Column<Date?>("Hire Date")
                .Column<int>("Salary")
                .Column<bool>("Is Active")
                .Column<BenchmarkStatus>("Status");
            reader.Convert(stream);
            if (sheet.Rows.Count != RowCount)
            {
                throw new($"Expected {RowCount} rows, got {sheet.Rows.Count}");
            }
        });

        Report("Dictionary", min, median, max);
    }
}
