[TestFixture]
public class BookReaderPrimitivesTests
{
    public enum SampleEnum
    {
        Alpha,
        Beta,
        Gamma
    }

    static async Task<IReadOnlyList<TModel>> RoundTrip<TModel>(IEnumerable<TModel> data)
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(data);
        await builder.ToStream(stream);
        stream.Position = 0;

        var reader = new BookReader();
        var sheet = reader.AddSheet<TModel>();
        reader.Convert(stream);
        return sheet.Rows;
    }

    [Test]
    public async Task Strings()
    {
        var rows = await RoundTrip(new[]
        {
            new StringRow { Value = "alpha" },
            new StringRow { Value = "beta" },
            new StringRow { Value = "" }
        });
        Assert.That(rows.Select(_ => _.Value), Is.EqualTo(new[] { "alpha", "beta", "" }));
    }

    public class StringRow
    {
        public string Value { get; set; } = "";
    }

    [Test]
    public async Task Booleans()
    {
        var rows = await RoundTrip(new[]
        {
            new BoolRow { Value = true, Nullable = true },
            new BoolRow { Value = false, Nullable = false },
            new BoolRow { Value = true, Nullable = null }
        });
        Assert.That(rows.Select(_ => _.Value), Is.EqualTo(new[] { true, false, true }));
        Assert.That(rows.Select(_ => _.Nullable), Is.EqualTo(new bool?[] { true, false, null }));
    }

    public class BoolRow
    {
        public bool Value { get; set; }
        public bool? Nullable { get; set; }
    }

    [Test]
    public async Task Integers()
    {
        var rows = await RoundTrip(new[]
        {
            new IntRow { Byte = 1, SByte = -1, Short = 2, UShort = 3, Int = 4, UInt = 5, Long = 6, ULong = 7 },
            new IntRow { Byte = 250, SByte = 100, Short = -32000, UShort = 60000, Int = -1000, UInt = 4_000_000_000, Long = -50_000_000_000, ULong = 9_000_000_000 }
        });

        Assert.That(rows[0].Byte, Is.EqualTo(1));
        Assert.That(rows[0].SByte, Is.EqualTo(-1));
        Assert.That(rows[0].Short, Is.EqualTo(2));
        Assert.That(rows[0].UShort, Is.EqualTo(3));
        Assert.That(rows[0].Int, Is.EqualTo(4));
        Assert.That(rows[0].UInt, Is.EqualTo(5));
        Assert.That(rows[0].Long, Is.EqualTo(6));
        Assert.That(rows[0].ULong, Is.EqualTo(7));

        Assert.That(rows[1].Byte, Is.EqualTo(250));
        Assert.That(rows[1].SByte, Is.EqualTo(100));
        Assert.That(rows[1].Short, Is.EqualTo(-32000));
        Assert.That(rows[1].UShort, Is.EqualTo(60000));
        Assert.That(rows[1].Int, Is.EqualTo(-1000));
        Assert.That(rows[1].UInt, Is.EqualTo(4_000_000_000));
        Assert.That(rows[1].Long, Is.EqualTo(-50_000_000_000));
        Assert.That(rows[1].ULong, Is.EqualTo(9_000_000_000));
    }

    public class IntRow
    {
        public byte Byte { get; set; }
        public sbyte SByte { get; set; }
        public short Short { get; set; }
        public ushort UShort { get; set; }
        public int Int { get; set; }
        public uint UInt { get; set; }
        public long Long { get; set; }
        public ulong ULong { get; set; }
    }

    [Test]
    public async Task Floats()
    {
        var rows = await RoundTrip(new[]
        {
            new FloatRow { Float = 1.5f, Double = 2.25, Decimal = 3.125m },
            new FloatRow { Float = -0.25f, Double = 0d, Decimal = -100m }
        });

        Assert.That(rows[0].Float, Is.EqualTo(1.5f));
        Assert.That(rows[0].Double, Is.EqualTo(2.25));
        Assert.That(rows[0].Decimal, Is.EqualTo(3.125m));
        Assert.That(rows[1].Float, Is.EqualTo(-0.25f));
        Assert.That(rows[1].Double, Is.EqualTo(0d));
        Assert.That(rows[1].Decimal, Is.EqualTo(-100m));
    }

    public class FloatRow
    {
        public float Float { get; set; }
        public double Double { get; set; }
        public decimal Decimal { get; set; }
    }

    [Test]
    public async Task NullableInts()
    {
        var rows = await RoundTrip(new[]
        {
            new NullableIntRow { Value = 42 },
            new NullableIntRow { Value = null }
        });
        Assert.That(rows.Select(_ => _.Value), Is.EqualTo(new int?[] { 42, null }));
    }

    public class NullableIntRow
    {
        public int? Value { get; set; }
    }

    [Test]
    public async Task DateTimes()
    {
        var rows = await RoundTrip(new[]
        {
            new DateTimeRow { Value = new(2020, 1, 15, 10, 30, 45) },
            new DateTimeRow { Value = new(1999, 12, 31, 23, 59, 59) }
        });
        Assert.That(rows[0].Value, Is.EqualTo(new DateTime(2020, 1, 15, 10, 30, 45)));
        Assert.That(rows[1].Value, Is.EqualTo(new DateTime(1999, 12, 31, 23, 59, 59)));
    }

    public class DateTimeRow
    {
        public DateTime Value { get; set; }
    }

    [Test]
    public async Task Dates()
    {
        var rows = await RoundTrip(new[]
        {
            new DateRow { Value = new(2020, 1, 15) },
            new DateRow { Value = new(2021, 7, 4) }
        });
        Assert.That(rows[0].Value, Is.EqualTo(new Date(2020, 1, 15)));
        Assert.That(rows[1].Value, Is.EqualTo(new Date(2021, 7, 4)));
    }

    public class DateRow
    {
        public Date Value { get; set; }
    }

    [Test]
    public async Task DateTimeOffsets()
    {
        var dto = new DateTimeOffset(2020, 5, 1, 12, 0, 0, TimeSpan.FromHours(0));
        var rows = await RoundTrip(new[]
        {
            new DateTimeOffsetRow { Value = dto }
        });
        Assert.That(rows[0].Value, Is.EqualTo(dto));
    }

    public class DateTimeOffsetRow
    {
        public DateTimeOffset Value { get; set; }
    }

    [Test]
    public async Task Guids()
    {
        var guid = Guid.Parse("11111111-2222-3333-4444-555555555555");
        var rows = await RoundTrip(new[]
        {
            new GuidRow { Value = guid }
        });
        Assert.That(rows[0].Value, Is.EqualTo(guid));
    }

    public class GuidRow
    {
        public Guid Value { get; set; }
    }

    [Test]
    public async Task Chars()
    {
        var rows = await RoundTrip(new[]
        {
            new CharRow { Value = 'A' },
            new CharRow { Value = 'z' }
        });
        Assert.That(rows[0].Value, Is.EqualTo('A'));
        Assert.That(rows[1].Value, Is.EqualTo('z'));
    }

    public class CharRow
    {
        public char Value { get; set; }
    }

    [Test]
    public async Task Enums()
    {
        var rows = await RoundTrip(new[]
        {
            new EnumRow { Value = SampleEnum.Alpha },
            new EnumRow { Value = SampleEnum.Beta },
            new EnumRow { Value = SampleEnum.Gamma }
        });
        Assert.That(rows.Select(_ => _.Value), Is.EqualTo(new[] { SampleEnum.Alpha, SampleEnum.Beta, SampleEnum.Gamma }));
    }

    public class EnumRow
    {
        public SampleEnum Value { get; set; }
    }

    [Test]
    public async Task NullableEnums()
    {
        var rows = await RoundTrip(new[]
        {
            new NullableEnumRow { Value = SampleEnum.Beta },
            new NullableEnumRow { Value = null }
        });
        Assert.That(rows.Select(_ => _.Value), Is.EqualTo(new SampleEnum?[] { SampleEnum.Beta, null }));
    }

    public class NullableEnumRow
    {
        public SampleEnum? Value { get; set; }
    }
}
