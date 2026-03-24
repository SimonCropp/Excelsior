namespace Excelsior.SourceGenerator;

[Generator]
public class SheetBuilderGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var results = context
            .SyntaxProvider
            .ForAttributeWithMetadataName(
                "Excelsior.SheetModelAttribute",
                (node, _) => node is TypeDeclarationSyntax,
                (context, _) => GetModelResult(context));

        context.RegisterSourceOutput(
            results,
            (productionContext, result) =>
            {
                if (result.Diagnostic is { } diagnostic)
                {
                    productionContext.ReportDiagnostic(
                        Diagnostic.Create(
                            InaccessibleTypeDescriptor,
                            diagnostic.Location,
                            diagnostic.TypeName));
                    return;
                }

                if (result.Model is not { } model)
                {
                    return;
                }

                var source = GenerateSource(model);
                productionContext.AddSource($"{model.TypeName}SheetBuilderExtensions.g.cs", source);
            });
    }

    static readonly DiagnosticDescriptor InaccessibleTypeDescriptor = new(
        "EXCEL002",
        "SheetModel type is not accessible",
        "Type '{0}' with [SheetModel] must be internal or public, including all containing types",
        "Excelsior",
        DiagnosticSeverity.Error,
        true);

    static ModelResult GetModelResult(GeneratorAttributeSyntaxContext context)
    {
        var type = (INamedTypeSymbol)context.TargetSymbol;

        if (!IsAccessible(type))
        {
            return new(
                null,
                new(type.Name, type.Locations.FirstOrDefault()));
        }

        var properties = new EquatableArray<PropertyInfo>(GetProperties(type).ToImmutableArray());

        if (properties.Length == 0)
        {
            return new(null, null);
        }

        var model = new ModelInfo(
            type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            GetFlattenedName(type),
            properties);
        return new(model, null);
    }

    static bool IsAccessible(INamedTypeSymbol type)
    {
        for (var current = type; current is not null; current = current.ContainingType)
        {
            if (current.DeclaredAccessibility is
                Accessibility.Private or
                Accessibility.Protected or
                Accessibility.ProtectedAndInternal)
            {
                return false;
            }
        }

        return true;
    }

    static string GetFlattenedName(INamedTypeSymbol type)
    {
        var parts = new List<string>();
        for (var current = type; current is not null; current = current.ContainingType)
        {
            parts.Add(current.Name);
        }

        parts.Reverse();
        return string.Concat(parts);
    }

    static IEnumerable<PropertyInfo> GetProperties(INamedTypeSymbol type) =>
        GetPropertiesRecursive(type, ImmutableArray<string>.Empty);

    static IEnumerable<PropertyInfo> GetPropertiesRecursive(INamedTypeSymbol type, ImmutableArray<string> parentPath)
    {
        foreach (var member in type.GetMembers())
        {
            if (member is not IPropertySymbol property)
            {
                continue;
            }

            if (property.DeclaredAccessibility != Accessibility.Public)
            {
                continue;
            }

            if (property.GetMethod is null)
            {
                continue;
            }

            if (HasAttribute(property, "Excelsior.IgnoreAttribute"))
            {
                continue;
            }

            if (HasAttribute(property, "Excelsior.SplitAttribute") ||
                HasAttribute(property.Type, "Excelsior.SplitAttribute"))
            {
                if (property.Type is INamedTypeSymbol namedType)
                {
                    var newPath = parentPath.Add(property.Name);
                    foreach (var nested in GetPropertiesRecursive(namedType, newPath))
                    {
                        yield return nested;
                    }
                }

                continue;
            }

            var propertyType = property.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            var accessPath = parentPath.IsEmpty
                ? property.Name
                : string.Join(".", parentPath.Add(property.Name));

            yield return new(property.Name, propertyType, accessPath);
        }
    }

    static bool HasAttribute(ISymbol symbol, string fullName) =>
        symbol.GetAttributes().Any(_ => _.AttributeClass?.ToDisplayString() == fullName);

    static bool HasAttribute(ITypeSymbol symbol, string fullName) =>
        symbol.GetAttributes().Any(_ => _.AttributeClass?.ToDisplayString() == fullName);

    static string GenerateSource(ModelInfo model)
    {
        var builder = new StringBuilder(
            $$"""
              // <auto-generated/>
              #nullable enable
              namespace Excelsior;
              using System;
              public static class {{model.TypeName}}SheetBuilderExtensions
              {

              """);

        for (var i = 0; i < model.Properties.Array.Length; i++)
        {
            if (i > 0)
            {
                builder.AppendLine();
            }

            var (prefix, propType, access) = model.Properties[i];
            var modelType = model.TypeFullName;

            builder.Append(
                $"""
                     public static ISheetBuilder<{modelType}, TStyle> {prefix}Column<TStyle>(
                         this ISheetBuilder<{modelType}, TStyle> builder,
                         Action<ColumnConfig<TStyle, {modelType}, {propType}>> configuration)
                         => builder.Column(_ => _.{access}, configuration);

                     public static void {prefix}HeadingText<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, string value)
                         => builder.HeadingText(_ => _.{access}, value);

                     public static void {prefix}Order<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, int? value)
                         => builder.Order(_ => _.{access}, value);

                     public static void {prefix}Width<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, int? value)
                         => builder.Width(_ => _.{access}, value);

                     public static void {prefix}HeadingStyle<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, Action<TStyle> value)
                         => builder.HeadingStyle(_ => _.{access}, value);

                     public static void {prefix}CellStyle<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, Action<TStyle, {modelType}, {propType}> value)
                         => builder.CellStyle(_ => _.{access}, value);

                     public static void {prefix}Format<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, string value)
                         => builder.Format(_ => _.{access}, value);

                     public static void {prefix}NullDisplay<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, string value)
                         => builder.NullDisplay(_ => _.{access}, value);

                     public static void {prefix}IsHtml<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder)
                         => builder.IsHtml(_ => _.{access});

                     public static void {prefix}Render<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, Func<{modelType}, {propType}, string?> value)
                         => builder.Render(_ => _.{access}, value);

                     public static void {prefix}Filter<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder)
                         => builder.Filter(_ => _.{access});

                     public static void {prefix}Include<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder, bool value)
                         => builder.Include(_ => _.{access}, value);

                     public static void {prefix}Exclude<TStyle>(this ISheetBuilder<{modelType}, TStyle> builder)
                         => builder.Exclude(_ => _.{access});

                 """);
        }

        builder.AppendLine("}");

        return builder.ToString();
    }
}
