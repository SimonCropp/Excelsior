using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

[TestFixture]
public class SheetBuilderExtensionsGeneratorTests
{
    [Test]
    public Task SimpleRecord()
    {
        var source = """
            using Excelsior;

            namespace TestModels;

            [ExcelsiorExtensions]
            public record Employee(string Email, int Age);
            """;

        return Verify(source);
    }

    [Test]
    public Task WithIgnore()
    {
        var source = """
            using Excelsior;

            namespace TestModels;

            [ExcelsiorExtensions]
            public record Employee(string Email, [Ignore] int Age, string Name);
            """;

        return Verify(source);
    }

    [Test]
    public Task WithSplitOnType()
    {
        var source = """
            using Excelsior;

            namespace TestModels;

            [Split]
            public record Address(string Street, string City);

            [ExcelsiorExtensions]
            public record Employee(string Email, Address Address);
            """;

        return Verify(source);
    }

    [Test]
    public Task WithSplitOnProperty()
    {
        var source = """
            using Excelsior;

            namespace TestModels;

            public record Address(string Street, string City);

            [ExcelsiorExtensions]
            public record Employee(string Email, [Split] Address Address);
            """;

        return Verify(source);
    }

    [Test]
    public Task ClassWithProperties()
    {
        var source = """
            using Excelsior;

            namespace TestModels;

            [ExcelsiorExtensions]
            public class Product
            {
                public string Name { get; set; }
                public decimal Price { get; set; }
            }
            """;

        return Verify(source);
    }

    static Task Verify(string source)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(source);

        var trustedAssemblies = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!)
            .Split(Path.PathSeparator)
            .Select(p => MetadataReference.CreateFromFile(p))
            .ToList();

        var excelsiorRef = MetadataReference.CreateFromFile(
            typeof(Excelsior.ExcelsiorExtensionsAttribute).Assembly.Location);

        var references = trustedAssemblies.Append(excelsiorRef);

        var compilation = CSharpCompilation.Create(
            "Tests",
            [syntaxTree],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        var generator = new Excelsior.SourceGenerator.SheetBuilderExtensionsGenerator();

        GeneratorDriver driver = CSharpGeneratorDriver.Create(generator);
        driver = driver.RunGenerators(compilation);

        var result = driver.GetRunResult();
        var generated = result.Results
            .SelectMany(r => r.GeneratedSources)
            .ToDictionary(s => s.HintName, s => s.SourceText.ToString());

        return Verifier.Verify(generated).UseDirectory("Snapshots");
    }
}
