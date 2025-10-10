namespace Excelsior;

[AttributeUsage(
    AttributeTargets.Property |
    AttributeTargets.Parameter |
    AttributeTargets.Class |
    AttributeTargets.Struct)]
public sealed class SplitAttribute :
    Attribute;