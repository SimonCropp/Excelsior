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
                            inaccessibleTypeDescriptor,
                            diagnostic.Location,
                            diagnostic.TypeName));
                    return;
                }

                if (result.Model is not { } model)
                {
                    return;
                }

                var source = GenerateExtensions(model);
                productionContext.AddSource($"{model.TypeName}SheetBuilderExtensions.g.cs", source);

                if (HasColumnAttributes(model))
                {
                    var registration = GenerateRegistration(model);
                    productionContext.AddSource($"{model.TypeName}ColumnAttributes.g.cs", registration);
                }
            });
    }

    static readonly DiagnosticDescriptor inaccessibleTypeDescriptor = new(
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

            var column = GetColumnData(property);
            var declaringType = property.ContainingType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);

            yield return new(property.Name, propertyType, accessPath, declaringType, column);
        }
    }

    static bool HasAttribute(ISymbol symbol, string fullName) =>
        symbol.GetAttributes().Any(_ => _.AttributeClass?.ToDisplayString() == fullName);

    static bool HasAttribute(ITypeSymbol symbol, string fullName) =>
        symbol.GetAttributes().Any(_ => _.AttributeClass?.ToDisplayString() == fullName);

    static AttributeData? FindColumnAttribute(IPropertySymbol property)
    {
        var attr = property.GetAttributes()
            .FirstOrDefault(_ => _.AttributeClass?.ToDisplayString() == "Excelsior.ColumnAttribute");

        if (attr is not null)
        {
            return attr;
        }

        // Check matching constructor parameter (e.g. record positional parameters)
        if (property.ContainingType is { } type)
        {
            foreach (var constructor in type.Constructors)
            {
                var param = constructor.Parameters
                    .FirstOrDefault(_ => _.Name == property.Name);

                if (param is null)
                {
                    continue;
                }

                attr = param.GetAttributes()
                    .FirstOrDefault(_ => _.AttributeClass?.ToDisplayString() == "Excelsior.ColumnAttribute");

                if (attr is not null)
                {
                    return attr;
                }
            }
        }

        return null;
    }

    static ColumnData? GetColumnData(IPropertySymbol property)
    {
        var attr = FindColumnAttribute(property);

        if (attr is null)
        {
            return null;
        }

        string? heading = null;
        int? order = null;
        int? width = null;
        string? format = null;
        string? nullDisplay = null;
        var isHtml = false;
        bool? filter = null;
        bool? include = null;

        foreach (var arg in attr.NamedArguments)
        {
            switch (arg.Key)
            {
                case "Heading":
                    heading = arg.Value.Value as string;
                    break;
                case "Order":
                    if (arg.Value.Value is int o and > -1)
                    {
                        order = o;
                    }

                    break;
                case "Width":
                    if (arg.Value.Value is int w and > -1)
                    {
                        width = w;
                    }

                    break;
                case "Format":
                    format = arg.Value.Value as string;
                    break;
                case "NullDisplay":
                    nullDisplay = arg.Value.Value as string;
                    break;
                case "IsHtml":
                    isHtml = arg.Value.Value is true;
                    break;
                case "Filter":
                    filter = arg.Value.Value as bool?;
                    break;
                case "Include":
                    include = arg.Value.Value as bool?;
                    break;
            }
        }

        return new(heading, order, width, format, nullDisplay, isHtml, filter, include);
    }

    static string GenerateExtensions(ModelInfo model)
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

            var prop = model.Properties[i];
            var prefix = prop.Name;
            var propType = prop.TypeFullName;
            var access = prop.AccessPath;
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

    static string GenerateRegistration(ModelInfo model)
    {
        var builder = new StringBuilder(
            $$"""
              // <auto-generated/>
              #nullable enable
              namespace Excelsior;
              using System.Runtime.CompilerServices;
              file static class {{model.TypeName}}ColumnAttributesRegistration
              {
                  [ModuleInitializer]
                  internal static void Register()
                  {

              """);

        foreach (var prop in model.Properties)
        {
            if (prop.Column is not { } col)
            {
                continue;
            }

            var args = new List<string>();

            if (col.Heading is { } heading)
            {
                args.Add($"Heading: \"{Escape(heading)}\"");
            }

            if (col.Order is { } order)
            {
                args.Add($"Order: {order}");
            }

            if (col.Width is { } width)
            {
                args.Add($"Width: {width}");
            }

            if (col.Format is { } format)
            {
                args.Add($"Format: \"{Escape(format)}\"");
            }

            if (col.NullDisplay is { } nullDisplay)
            {
                args.Add($"NullDisplay: \"{Escape(nullDisplay)}\"");
            }

            if (col.IsHtml)
            {
                args.Add("IsHtml: true");
            }

            if (col.Filter is { } filter)
            {
                args.Add($"Filter: {(filter ? "true" : "false")}");
            }

            if (col.Include is { } include)
            {
                args.Add($"Include: {(include ? "true" : "false")}");
            }

            builder.AppendLine(
                $"        GeneratedColumnAttributes.Register(typeof({prop.DeclaringTypeFullName}), \"{prop.Name}\", new({string.Join(", ", args)}));");
        }

        builder.Append(
            """
                }
            }
            """);

        return builder.ToString();
    }

    static bool HasColumnAttributes(ModelInfo model) =>
        model.Properties.Any(_ => _.Column is not null);

    static string Escape(string value) =>
        value.Replace("\\", "\\\\").Replace("\"", "\\\"");
}
