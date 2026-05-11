namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Parameter)]
public sealed class IgnoreAttribute :
    Attribute;