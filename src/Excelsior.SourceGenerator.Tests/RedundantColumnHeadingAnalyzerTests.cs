using System.Collections.Immutable;
using Microsoft.CodeAnalysis.Diagnostics;

[TestFixture]
public class RedundantColumnHeadingAnalyzerTests
{
    [Test]
    public void RedundantHeading_OnProperty()
    {
        var source = """
            using Excelsior;

            public class Order
            {
                [Column(Heading = "ReferenceNumber")]
                public string ReferenceNumber { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(1, diagnostics.Length);
        AreEqual("EXCEL001", diagnostics[0].Id);
        IsTrue(diagnostics[0].GetMessage().Contains("ReferenceNumber"));
    }

    [Test]
    public void RedundantHeading_OnRecordParameter()
    {
        var source = """
            using Excelsior;

            public record Order([Column(Heading = "Name")] string Name);
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(1, diagnostics.Length);
        AreEqual("EXCEL001", diagnostics[0].Id);
    }

    [Test]
    public void DifferentHeading_NoDiagnostic()
    {
        var source = """
            using Excelsior;

            public class Order
            {
                [Column(Heading = "Reference Number")]
                public string ReferenceNumber { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(0, diagnostics.Length);
    }

    [Test]
    public void NoHeading_NoDiagnostic()
    {
        var source = """
            using Excelsior;

            public class Order
            {
                [Column(Width = 15)]
                public string ReferenceNumber { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(0, diagnostics.Length);
    }

    [Test]
    public void NoColumnAttribute_NoDiagnostic()
    {
        var source = """
            public class Order
            {
                public string ReferenceNumber { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(0, diagnostics.Length);
    }

    static ImmutableArray<Diagnostic> GetDiagnostics(string source)
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

        var analyzer = new Excelsior.SourceGenerator.RedundantColumnHeadingAnalyzer();

        return compilation
            .WithAnalyzers([analyzer])
            .GetAnalyzerDiagnosticsAsync()
            .GetAwaiter()
            .GetResult();
    }
}
