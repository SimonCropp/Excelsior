namespace Excelsior;

class ColumnReadState<TModel>
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