using System.Collections.Immutable;
using Microsoft.CodeAnalysis.Diagnostics;

[TestFixture]
public class MismatchedIsHtmlAnalyzerTests
{
    [Test]
    public void Mismatch_OnProperty()
    {
        var source = """
            using Excelsior;
            using System.Diagnostics.CodeAnalysis;

            public class Order
            {
                [Column(IsHtml = false)]
                [StringSyntax("html")]
                public string Notes { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(1, diagnostics.Length);
        AreEqual("EXCEL003", diagnostics[0].Id);
    }

    [Test]
    public void Mismatch_OnRecordParameter()
    {
        var source = """
            using Excelsior;
            using System.Diagnostics.CodeAnalysis;

            public record Order([Column(IsHtml = false), StringSyntax("html")] string Notes);
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(1, diagnostics.Length);
        AreEqual("EXCEL003", diagnostics[0].Id);
    }

    [Test]
    public void Mismatch_CaseInsensitive()
    {
        var source = """
            using Excelsior;
            using System.Diagnostics.CodeAnalysis;

            public class Order
            {
                [Column(IsHtml = false)]
                [StringSyntax("HTML")]
                public string Notes { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(1, diagnostics.Length);
        AreEqual("EXCEL003", diagnostics[0].Id);
    }

    [Test]
    public void ColumnTrueWithStringSyntax_NoDiagnostic()
    {
        var source = """
            using Excelsior;
            using System.Diagnostics.CodeAnalysis;

            public class Order
            {
                [Column(IsHtml = true)]
                [StringSyntax("html")]
                public string Notes { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(0, diagnostics.Length);
    }

    [Test]
    public void OnlyStringSyntax_NoDiagnostic()
    {
        var source = """
            using System.Diagnostics.CodeAnalysis;

            public class Order
            {
                [StringSyntax("html")]
                public string Notes { get; set; }
            }
            """;

        var diagnostics = GetDiagnostics(source);

        AreEqual(0, diagnostics.Length);
    }

    [Test]
    public void NonHtmlStringSyntax_NoDiagnostic()
    {
        var source = """
            using Excelsior;
            using System.Diagnostics.CodeAnalysis;

            public class Order
            {
                [Column(IsHtml = false)]
                [StringSyntax("json")]
                public string Notes { get; set; }
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

        var analyzer = new Excelsior.SourceGenerator.MismatchedIsHtmlAnalyzer();

        return compilation
            .WithAnalyzers([analyzer])
            .GetAnalyzerDiagnosticsAsync()
            .GetAwaiter()
            .GetResult();
    }
}
