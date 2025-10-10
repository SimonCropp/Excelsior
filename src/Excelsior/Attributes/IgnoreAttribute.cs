namespace Excelsior;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter)]
public sealed class IgnoreAttribute :
    Attribute;