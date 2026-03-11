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

    static Dictionary<string, string> Generate(string source)
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

        var result = driver.GetRunResult();
        return result.Results
            .SelectMany(_ => _.GeneratedSources)
            .ToDictionary(
                _ => _.HintName,
                _ => _.SourceText.ToString());
    }
}
