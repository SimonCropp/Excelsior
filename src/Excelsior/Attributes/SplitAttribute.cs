namespace Excelsior;

[AttributeUsage(
    AttributeTargets.Property |
    AttributeTargets.Parameter |
    AttributeTargets.Class |
    AttributeTargets.Struct)]
public sealed class SplitAttribute(bool useHierachyForName = false) :
    Attribute
{
    public bool UseHierachyForName { get; } = useHierachyForName;
}