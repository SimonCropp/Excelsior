namespace Excelsior.SourceGenerator;

record struct MemberAccess(
    ISymbol Symbol,
    ITypeSymbol Type,
    bool HasGetter,
    bool HasPublicSetter,
    bool IsInitOnly,
    bool IsRequired)
{
    public string Name => Symbol.Name;
}

static class MemberAccessExtensions
{
    /// <summary>
    /// Yields the public instance properties and fields that participate in column
    /// discovery — mirroring the runtime <c>Properties&lt;T&gt;</c> / setter map.
    /// </summary>
    public static IEnumerable<MemberAccess> EnumerateColumnMembers(this INamedTypeSymbol type)
    {
        foreach (var member in type.GetMembers())
        {
            if (member.DeclaredAccessibility != Accessibility.Public)
            {
                continue;
            }

            if (member.IsStatic)
            {
                continue;
            }

            switch (member)
            {
                case IPropertySymbol { IsIndexer: false, GetMethod: not null } property:
                    yield return new(
                        property,
                        property.Type,
                        HasGetter: true,
                        HasPublicSetter: property.SetMethod is { DeclaredAccessibility: Accessibility.Public },
                        IsInitOnly: property.SetMethod?.IsInitOnly == true,
                        IsRequired: property.IsRequired);
                    break;

                case IFieldSymbol { IsConst: false, IsImplicitlyDeclared: false } field:
                    // readonly fields cannot be assigned by an external activator/setter,
                    // so we expose them as readable-only (HasPublicSetter: false). They
                    // still participate in discovery for the write path.
                    yield return new(
                        field,
                        field.Type,
                        HasGetter: true,
                        HasPublicSetter: !field.IsReadOnly,
                        IsInitOnly: false,
                        IsRequired: field.IsRequired);
                    break;
            }
        }
    }
}
