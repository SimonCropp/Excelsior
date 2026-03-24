[TestFixture]
public class GeneratorTests
{
    [Test]
    public Task SimpleRecord()
    {
        var source = """
            using Excelsior;

            [SheetModel]
            public record Employee(string Email, int Age);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithIgnore()
    {
        var source = """
            using Excelsior;

            [SheetModel]
            public record Employee(string Email, [Ignore] int Age, string Name);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithSplitOnType()
    {
        var source = """
            using Excelsior;

            [Split]
            public record Address(string Street, string City);

            [SheetModel]
            public record Employee(string Email, Address Address);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithSplitOnProperty()
    {
        var source = """
            using Excelsior;

            public record Address(string Street, string City);

            [SheetModel]
            public record Employee(string Email, [Split] Address Address);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task NestedUnderSameParent()
    {
        var source = """
            using Excelsior;

            public class Models
            {
                [SheetModel]
                public record Employee(string Email, int Age);

                [SheetModel]
                public record Product(string Name, decimal Price);
            }
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task ClassWithProperties()
    {
        var source = """
            using Excelsior;

            [SheetModel]
            public class Product
            {
                public string Name { get; set; }
                public decimal Price { get; set; }
            }
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithColumnAttributes()
    {
        var source = """
            using Excelsior;

            [SheetModel]
            public class Employee
            {
                [Column(Heading = "Employee ID", Order = 1, Format = "0000")]
                public int Id { get; set; }

                [Column(Heading = "Full Name", Order = 2, Width = 20)]
                public string Name { get; set; }

                [Column(Heading = "Email Address", Width = 30)]
                public string Email { get; set; }

                [Column(Order = 3, NullDisplay = "unknown")]
                public string HireDate { get; set; }

                [Column(IsHtml = true)]
                public string Notes { get; set; }

                [Column(Filter = true)]
                public string Department { get; set; }

                [Column(Include = false)]
                public string Internal { get; set; }

                public int Age { get; set; }
            }
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithColumnAttributesOnRecordParameters()
    {
        var source = """
            using Excelsior;

            [SheetModel]
            public record Employee(
                [Column(Heading = "Employee ID", Order = 1)] int Id,
                [Column(Width = 20)] string Name);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task WithColumnAttributesOnSplitProperties()
    {
        var source = """
            using Excelsior;

            [Split]
            public record Address(
                [Column(Heading = "Street Address", Width = 30)] string Street,
                string City);

            [SheetModel]
            public record Employee(string Name, Address Address);
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public Task NestedUnderDifferentParentsWithSameName()
    {
        var source = """
            using Excelsior;

            public class GroupA
            {
                [SheetModel]
                public record Row(string Name, int Age);
            }

            public class GroupB
            {
                [SheetModel]
                public record Row(string Email, decimal Salary);
            }
            """;

        var generated = Generate(source);

        return Verify(generated);
    }

    [Test]
    public void PrivateNestedTypeProducesError()
    {
        var source = """
            using Excelsior;

            public static class Container
            {
                [SheetModel]
                record Row(string Name, int Age);
            }
            """;

        var result = RunGenerator(source);
        var diagnostics = result.Diagnostics;

        Assert.That(result.Results.SelectMany(_ => _.GeneratedSources), Is.Empty);
        Assert.That(diagnostics, Has.Length.EqualTo(1));
        Assert.That(diagnostics[0].Id, Is.EqualTo("EXCEL002"));
        Assert.That(diagnostics[0].Severity, Is.EqualTo(DiagnosticSeverity.Error));
    }

    static Dictionary<string, string> Generate(string source)
    {
        var result = RunGenerator(source);
        return result.Results
            .SelectMany(_ => _.GeneratedSources)
            .ToDictionary(
                _ => _.HintName,
                _ => _.SourceText.ToString());
    }

    static GeneratorDriverRunResult RunGenerator(string source)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        var trustedAssemblies = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!)
            .Split(Path.PathSeparator)
            .Select(_ => MetadataReference.CreateFromFile(_))
            .ToList();

        var excelsiorRef = MetadataReference.CreateFromFile(
            typeof(Excelsior.SheetModelAttribute).Assembly.Location);

        var references = trustedAssemblies.Append(excelsiorRef);

        var compilation = CSharpCompilation.Create(
            "Tests",
            [syntaxTree],
            references,
            new(OutputKind.DynamicallyLinkedLibrary));

        var generator = new Excelsior.SourceGenerator.SheetBuilderGenerator();

        GeneratorDriver driver = CSharpGeneratorDriver.Create(generator);
        driver = driver.RunGenerators(compilation);

        return driver.GetRunResult();
    }
}
