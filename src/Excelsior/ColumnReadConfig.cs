namespace Excelsior;

class ColumnReadConfig<TModel>
{
    public required string Name { get; init; }
    public required string Heading { get; set; }
    public required Type Type { get; init; }
    public required int? Order { get; set; }
    public required int DeclarationIndex { get; init; }
    public required bool Include { get; set; }
    public required Action<TModel, object?> SetValue { get; init; }
    public Func<Cell, object?>? Convert { get; set; }
}

public class ColumnReadConfig<TModel, TProperty>
{
    public string? Heading { get; set; }
    public int? Order { get; set; }
    public bool? Include { get; set; }

    /// <summary>
    /// Custom conversion delegate. Receives the underlying OpenXml <see cref="Cell"/>
    /// and returns the value to assign to the property. When set, this takes
    /// precedence over the built-in cell parsing.
    /// </summary>
    public Func<Cell, TProperty>? Convert { get; set; }
}
