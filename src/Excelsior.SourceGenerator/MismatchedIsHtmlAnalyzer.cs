namespace Excelsior.SourceGenerator;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public class MismatchedIsHtmlAnalyzer : DiagnosticAnalyzer
{
    public static readonly DiagnosticDescriptor Rule = new(
        "EXCEL003",
        "Mismatched IsHtml",
        "[Column(IsHtml = false)] contradicts [StringSyntax(\"html\")] on the same member",
        "Excelsior.Usage",
        DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [Rule];

    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
        context.RegisterSymbolAction(AnalyzeParameter, SymbolKind.Parameter);
    }

    static void AnalyzeProperty(SymbolAnalysisContext context) =>
        AnalyzeSymbol(context, context.Symbol.GetAttributes());

    static void AnalyzeParameter(SymbolAnalysisContext context) =>
        AnalyzeSymbol(context, context.Symbol.GetAttributes());

    static void AnalyzeSymbol(SymbolAnalysisContext context, ImmutableArray<AttributeData> attributes)
    {
        AttributeData? columnAttribute = null;
        var explicitIsHtmlFalse = false;
        var hasHtmlSyntax = false;
        AttributeData? htmlSyntaxAttribute = null;

        foreach (var attribute in attributes)
        {
            var attrClass = attribute.AttributeClass;
            if (attrClass is null)
            {
                continue;
            }

            if (attrClass.Name == "ColumnAttribute" &&
                attrClass.ContainingNamespace?.ToDisplayString() == "Excelsior")
            {
                columnAttribute = attribute;
                foreach (var named in attribute.NamedArguments)
                {
                    if (named.Key == "IsHtml" && named.Value.Value is false)
                    {
                        explicitIsHtmlFalse = true;
                    }
                }

                continue;
            }

            if (attrClass.Name == "StringSyntaxAttribute" &&
                attrClass.ContainingNamespace?.ToDisplayString() == "System.Diagnostics.CodeAnalysis" &&
                attribute.ConstructorArguments.Length > 0 &&
                attribute.ConstructorArguments[0].Value is string syntax &&
                string.Equals(syntax, "html", System.StringComparison.OrdinalIgnoreCase))
            {
                hasHtmlSyntax = true;
                htmlSyntaxAttribute = attribute;
            }
        }

        if (!explicitIsHtmlFalse || !hasHtmlSyntax)
        {
            return;
        }

        var location = htmlSyntaxAttribute?.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                       ?? columnAttribute?.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation();

        if (location is null)
        {
            return;
        }

        context.ReportDiagnostic(Diagnostic.Create(Rule, location));
    }
}
