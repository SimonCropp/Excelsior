record struct ModelInfo(
    string TypeFullName,
    string TypeName,
    EquatableArray<PropertyInfo> Properties,
    ActivatorPlan? Activator);

record struct ActivatorPlan(
    EquatableArray<ActivatorParam> CtorParams,
    EquatableArray<ActivatorAssign> InitProps,
    EquatableArray<ActivatorAssign> SetProps);

record struct ActivatorParam(
    string Name,
    string TypeFullName);

record struct ActivatorAssign(
    string Name,
    string TypeFullName);

record struct ModelResult(
    ModelInfo? Model,
    DiagnosticInfo? Diagnostic);

record struct DiagnosticInfo(
    string TypeName,
    Location? Location);
