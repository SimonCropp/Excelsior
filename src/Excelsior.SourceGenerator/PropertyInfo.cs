record struct PropertyInfo(
    string Name,
    string TypeFullName,
    string AccessPath,
    string DeclaringTypeFullName,
    ColumnData? Column = null);

record struct ColumnData(
    string? Heading = null,
    int? Order = null,
    int? Width = null,
    string? Format = null,
    string? NullDisplay = null,
    bool IsHtml = false,
    bool? Filter = null,
    bool? Include = null);
