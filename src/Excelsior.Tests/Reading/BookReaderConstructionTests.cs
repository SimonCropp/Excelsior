// ReSharper disable NotAccessedPositionalProperty.Local
[TestFixture]
public class BookReaderConstructionTests
{
    static async Task<MemoryStream> Write<T>(params IEnumerable<T> rows)
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(rows);
        await builder.ToStream(stream);
        stream.Position = 0;
        return stream;
    }

    #region BookReaderPositionalRecord

    public record PersonRecord(string Name, int Age);

    [Test]
    public async Task PositionalRecord()
    {
        var stream = await Write(
            new PersonRecord("Alice", 30),
            new PersonRecord("Bob", 25));

        var reader = new BookReader();
        var sheet = reader.AddSheet<PersonRecord>();
        reader.Convert(stream);

        Assert.That(
            sheet.Rows,
            Is.EqualTo<PersonRecord>(
            [
                new("Alice", 30),
                new("Bob", 25)
            ]));
    }

    #endregion

    public class CtorByName
    {
        public string Name { get; }
        public int Age { get; }

#pragma warning disable IDE0290
        // ReSharper disable InconsistentNaming
        public CtorByName(string Name, int Age)
            // ReSharper restore InconsistentNaming
#pragma warning restore IDE0290
        {
            this.Name = Name;
            this.Age = Age;
        }
    }

    [Test]
    public async Task ConstructorByName()
    {
        var stream = await Write(
            new CtorByName("Alice", 30),
            new CtorByName("Bob", 25));

        var reader = new BookReader();
        var sheet = reader.AddSheet<CtorByName>();
        reader.Convert(stream);

        Assert.That(
            sheet.Rows.Select(_ => _.Name),
            Is.EqualTo([
                "Alice",
                "Bob"
            ]));
        Assert.That(sheet.Rows.Select(_ => _.Age),
            Is.EqualTo([
                30,
                25
            ]));
    }

    public class ParameterlessWins
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }

        public ParameterlessWins()
        {
        }

        // ReSharper disable InconsistentNaming
        // ReSharper disable UnusedParameter.Local
        public ParameterlessWins(string Name, int Age) =>
            // ReSharper restore UnusedParameter.Local
            // ReSharper restore InconsistentNaming
            throw new("Parameterless ctor should be preferred when present");
    }

    [Test]
    public async Task ParameterlessCtorWinsOverMatching()
    {
        var stream = await Write(
            new ParameterlessWins
            {
                Name = "Alice",
                Age = 30
            });

        var reader = new BookReader();
        var sheet = reader.AddSheet<ParameterlessWins>();
        reader.Convert(stream);

        Assert.That(sheet.Rows[0].Name, Is.EqualTo("Alice"));
        Assert.That(sheet.Rows[0].Age, Is.EqualTo(30));
    }

    public class PrivateParameterless
    {
        public string Name { get; set; } = "";

        // ReSharper disable once UnusedMember.Local
        PrivateParameterless()
        {
        }

        public PrivateParameterless(string name) =>
            Name = name;
    }

    [Test]
    public async Task NonPublicParameterlessIsUsed()
    {
        var stream = await Write(new PrivateParameterless("Alice"));

        var reader = new BookReader();
        var sheet = reader.AddSheet<PrivateParameterless>();
        reader.Convert(stream);

        Assert.That(sheet.Rows[0].Name, Is.EqualTo("Alice"));
    }

    public class PartialCtor
    {
        public string Name { get; }
        public string? Notes { get; set; }

        // ReSharper disable once ConvertToPrimaryConstructor
        // ReSharper disable once InconsistentNaming
        public PartialCtor(string Name) =>
            this.Name = Name;
    }

    [Test]
    public async Task ConstructorParamsAndSettersCombine()
    {
        var stream = await Write(new PartialCtor("Alice")
        {
            Notes = "VIP"
        });

        var reader = new BookReader();
        var sheet = reader.AddSheet<PartialCtor>();
        reader.Convert(stream);

        Assert.That(sheet.Rows[0].Name, Is.EqualTo("Alice"));
        Assert.That(sheet.Rows[0].Notes, Is.EqualTo("VIP"));
    }

    public class CaseMismatch(string name)
    {
        public string Name { get; } = name;
    }

    [Test]
    public async Task ConstructorParamMatchIsCaseSensitive()
    {
        var stream = await Write(new CaseMismatch("Alice"));

        var reader = new BookReader();
        var sheet = reader.AddSheet<CaseMismatch>();
        reader.Convert(stream);

        // Property is "Name"; ctor param is "name". Lookup is case-sensitive,
        // so the ctor receives the default value (null).
        Assert.That(sheet.Rows[0].Name, Is.Null);
    }

    public class HasName
    {
        public required string Name { get; init; }
    }

    public class NoUsableCtor
    {
        public string Name { get; set; } = "";

        internal NoUsableCtor(string name) =>
            Name = name;
    }

    [Test]
    public async Task NoUsableCtorThrows()
    {
        // Write a stream with column "Name" using a writable model, then try
        // to read it into a model that has no public parameterless ctor and
        // no public ctor at all (only internal).
        var stream = await Write(
            new HasName
            {
                Name = "Alice"
            });

        var reader = new BookReader();
        reader.AddSheet<NoUsableCtor>();

        var exception = Assert.Throws<Exception>(() => reader.Convert(stream))!;
        Assert.That(exception.Message, Does.Contain("no usable constructor"));
    }
}
