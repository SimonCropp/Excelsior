namespace Excelsior.SourceGenerator;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public class RedundantColumnHeadingAnalyzer : DiagnosticAnalyzer
{
    public static readonly DiagnosticDescriptor Rule = new(
        "EXCEL001",
        "Redundant Column Heading",
        "Heading \"{0}\" matches the property name; remove it to use the default heading",
        "Excelsior.Usage",
        DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [Rule];

    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();
        context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
        context.RegisterSymbolAction(AnalyzeParameter, SymbolKind.Parameter);
    }

    static void AnalyzeProperty(SymbolAnalysisContext context)
    {
        var property = (IPropertySymbol)context.Symbol;
        AnalyzeSymbol(context, property.Name);
    }

    static void AnalyzeParameter(SymbolAnalysisContext context)
    {
        var parameter = (IParameterSymbol)context.Symbol;
        AnalyzeSymbol(context, parameter.Name);
    }

    static void AnalyzeSymbol(SymbolAnalysisContext context, string memberName)
    {
        foreach (var attribute in context.Symbol.GetAttributes())
        {
            var attrClass = attribute.AttributeClass;
            if (attrClass is null)
            {
                continue;
            }

            if (attrClass.Name != "ColumnAttribute" ||
                attrClass.ContainingNamespace?.ToDisplayString() != "Excelsior")
            {
                continue;
            }

            foreach (var namedArg in attribute.NamedArguments)
            {
                if (namedArg.Key != "Heading")
                {
                    continue;
                }

                if (namedArg.Value.Value is not string heading)
                {
                    continue;
                }

                if (!string.Equals(heading, memberName, StringComparison.Ordinal) &&
                    !string.Equals(heading, CamelCase.Split(memberName), StringComparison.Ordinal))
                {
                    continue;
                }

                var location = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation();
                if (location is null)
                {
                    continue;
                }

                context.ReportDiagnostic(Diagnostic.Create(Rule, location, heading));
            }
        }
    }

}
