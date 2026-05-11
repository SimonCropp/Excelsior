using JetBrains.Annotations;

namespace Excelsior;

[AttributeUsage(
    AttributeTargets.Class |
    AttributeTargets.Struct)]
[MeansImplicitUse(ImplicitUseTargetFlags.Members)]
public sealed class SheetModelAttribute :
    Attribute;
