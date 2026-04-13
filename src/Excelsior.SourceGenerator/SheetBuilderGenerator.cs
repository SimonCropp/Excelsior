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

            if (HasAttribute(property, "IgnoreAttribute"))
            {
                continue;
            }

            if (HasAttribute(property, "SplitAttribute") ||
                HasAttribute(property.Type, "SplitAttribute"))
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

            var propertyType = property.Type.ToDisplayString(
                SymbolDisplayFormat.FullyQualifiedFormat
                    .AddMiscellaneousOptions(SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier));
            var accessPath = parentPath.IsEmpty
                ? property.Name
                : string.Join(".", parentPath.Add(property.Name));

            var column = GetColumnData(property);
            var declaringType = property.ContainingType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);

            yield return new(property.Name, propertyType, accessPath, declaringType, column);
        }
    }

    static bool IsAttribute(INamedTypeSymbol? type, string name) =>
        type is { Name: var n, ContainingNamespace: { Name: "Excelsior", ContainingNamespace.IsGlobalNamespace: true } }
        && n == name;

    static bool HasAttribute(ISymbol symbol, string name) =>
        symbol.GetAttributes().Any(_ => IsAttribute(_.AttributeClass, name));

    static bool HasAttribute(ITypeSymbol symbol, string name) =>
        symbol.GetAttributes().Any(_ => IsAttribute(_.AttributeClass, name));

    static AttributeData? FindAttribute(ImmutableArray<AttributeData> attributes, string name) =>
        attributes.FirstOrDefault(_ => IsAttribute(_.AttributeClass, name));

    static AttributeData? FindColumnAttribute(IPropertySymbol property)
    {
        var attr = FindAttribute(property.GetAttributes(), "ColumnAttribute");

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

                attr = FindAttribute(param.GetAttributes(), "ColumnAttribute");

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
        int? minWidth = null;
        int? maxWidth = null;
        string? format = null;
        string? nullDisplay = null;
        var isHtml = false;
        bool? filter = null;
        bool? include = null;

        foreach (var arg in attr.NamedArguments)
        {
            var value = arg.Value.Value;
            switch (arg.Key)
            {
                case "Heading":
                    heading = value as string;
                    break;
                case "Order":
                    if (value is int o and > -1)
                    {
                        order = o;
                    }

                    break;
                case "Width":
                    if (value is int w and > -1)
                    {
                        width = w;
                    }

                    break;
                case "MinWidth":
                    if (value is int minW and > -1)
                    {
                        minWidth = minW;
                    }

                    break;
                case "MaxWidth":
                    if (value is int maxW and > -1)
                    {
                        maxWidth = maxW;
                    }

                    break;
                case "Format":
                    format = value as string;
                    break;
                case "NullDisplay":
                    nullDisplay = value as string;
                    break;
                case "IsHtml":
                    isHtml = value is true;
                    break;
                case "Filter":
                    filter = value as bool?;
                    break;
                case "Include":
                    include = value as bool?;
                    break;
            }
        }

        return new(heading, order, width, minWidth, maxWidth, format, nullDisplay, isHtml, filter, include);
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
                     public static ISheetBuilder<{modelType}> {prefix}Column(
                         this ISheetBuilder<{modelType}> builder,
                         Action<ColumnConfig<{modelType}, {propType}>> configuration)
                         => builder.Column(_ => _.{access}, configuration);

                     public static void {prefix}HeadingText(this ISheetBuilder<{modelType}> builder, string value)
                         => builder.HeadingText(_ => _.{access}, value);

                     public static void {prefix}Order(this ISheetBuilder<{modelType}> builder, int? value)
                         => builder.Order(_ => _.{access}, value);

                     public static void {prefix}Width(this ISheetBuilder<{modelType}> builder, int? value)
                         => builder.Width(_ => _.{access}, value);

                     public static void {prefix}MinWidth(this ISheetBuilder<{modelType}> builder, int? value)
                         => builder.MinWidth(_ => _.{access}, value);

                     public static void {prefix}MaxWidth(this ISheetBuilder<{modelType}> builder, int? value)
                         => builder.MaxWidth(_ => _.{access}, value);

                     public static void {prefix}HeadingStyle(this ISheetBuilder<{modelType}> builder, Action<Excelsior.CellStyle> value)
                         => builder.HeadingStyle(_ => _.{access}, value);

                     public static void {prefix}CellStyle(this ISheetBuilder<{modelType}> builder, Action<Excelsior.CellStyle, {modelType}, {propType}> value)
                         => builder.CellStyle(_ => _.{access}, value);

                     public static void {prefix}Format(this ISheetBuilder<{modelType}> builder, string value)
                         => builder.Format(_ => _.{access}, value);

                     public static void {prefix}NullDisplay(this ISheetBuilder<{modelType}> builder, string value)
                         => builder.NullDisplay(_ => _.{access}, value);

                     public static void {prefix}IsHtml(this ISheetBuilder<{modelType}> builder)
                         => builder.IsHtml(_ => _.{access});

                     public static void {prefix}Render(this ISheetBuilder<{modelType}> builder, Func<{modelType}, {propType}, string?> value)
                         => builder.Render(_ => _.{access}, value);

                     public static void {prefix}Formula(this ISheetBuilder<{modelType}> builder, Func<{modelType}, FormulaContext<{modelType}>, string> value)
                         => builder.Formula(_ => _.{access}, value);

                     public static void {prefix}Formula(this ISheetBuilder<{modelType}> builder, Func<FormulaContext<{modelType}>, string> value)
                         => builder.Formula(_ => _.{access}, value);

                     public static void {prefix}Filter(this ISheetBuilder<{modelType}> builder)
                         => builder.Filter(_ => _.{access});

                     public static void {prefix}Include(this ISheetBuilder<{modelType}> builder, bool value)
                         => builder.Include(_ => _.{access}, value);

                     public static void {prefix}Exclude(this ISheetBuilder<{modelType}> builder)
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

            if (col.MinWidth is { } minWidth)
            {
                args.Add($"MinWidth: {minWidth}");
            }

            if (col.MaxWidth is { } maxWidth)
            {
                args.Add($"MaxWidth: {maxWidth}");
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
