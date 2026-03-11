namespace Excelsior.SourceGenerator;

record struct ModelInfo(
    string TypeFullName,
    string TypeName,
    EquatableArray<PropertyInfo> Properties);