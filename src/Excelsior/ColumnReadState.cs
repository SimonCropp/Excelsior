namespace Excelsior;

class ColumnReadState
{
    public required string Name { get; init; }
    public required string Heading { get; set; }
    public required Type Type { get; init; }
    public required bool Include { get; set; }
    public Func<Cell, object?>? Convert { get; set; }
}