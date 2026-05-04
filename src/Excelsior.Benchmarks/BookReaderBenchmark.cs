using BenchmarkDotNet.Attributes;

namespace Excelsior.Benchmarks;

[MemoryDiagnoser]
public class BookReaderBenchmark
{
    public const int RowCount = 10_000;

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

    byte[] reflectionWorkbook = null!;
    byte[] sourceGenWorkbook = null!;

    [GlobalSetup]
    public async Task Setup()
    {
        reflectionWorkbook = await WriteWorkbook(GenerateReflection());
        sourceGenWorkbook = await WriteWorkbook(GenerateSourceGen());
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

    [Benchmark]
    public int TypedReflection()
    {
        using var stream = new MemoryStream(reflectionWorkbook, writable: false);
        var reader = new BookReader();
        var sheet = reader.AddSheet<ReflectionModel>();
        reader.Convert(stream);
        return sheet.Rows.Count;
    }

    [Benchmark]
    public int TypedSourceGen()
    {
        using var stream = new MemoryStream(sourceGenWorkbook, writable: false);
        var reader = new BookReader();
        var sheet = reader.AddSheet<SourceGenModel>();
        reader.Convert(stream);
        return sheet.Rows.Count;
    }

    [Benchmark]
    public int Dictionary()
    {
        using var stream = new MemoryStream(reflectionWorkbook, writable: false);
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
        return sheet.Rows.Count;
    }
}
