// Internal stubs of JetBrains.Annotations attributes. ReSharper / Rider recognise these by
// namespace + name match against metadata, regardless of which assembly they live in — so
// shipping our own internal copy lets us decorate ColumnAttribute with [MeansImplicitUse]
// without taking a dependency on the JetBrains.Annotations NuGet package. Users may freely
// also reference that package; the internal scope here prevents type-resolution conflicts.
namespace JetBrains.Annotations;

[AttributeUsage(AttributeTargets.Class)]
sealed class MeansImplicitUseAttribute(
    ImplicitUseTargetFlags targetFlags = ImplicitUseTargetFlags.Default) :
    Attribute
{
    public ImplicitUseTargetFlags TargetFlags { get; } = targetFlags;
}

[Flags]
enum ImplicitUseTargetFlags
{
    Default = Itself,
    Itself = 1,
    Members = 2,
    WithInheritors = 4,
    WithMembers = Itself | Members,
}
