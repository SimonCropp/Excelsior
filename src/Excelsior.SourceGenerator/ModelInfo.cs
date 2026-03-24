record struct ModelInfo(
    string TypeFullName,
    string TypeName,
    EquatableArray<PropertyInfo> Properties);

record struct ModelResult(
    ModelInfo? Model,
    DiagnosticInfo? Diagnostic);

record struct DiagnosticInfo(
    string TypeName,
    Location? Location);
