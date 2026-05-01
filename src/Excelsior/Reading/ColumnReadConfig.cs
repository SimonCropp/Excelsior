namespace Excelsior;

public class ColumnReadConfig<TProperty>
{
    public string? Heading { get; set; }

    /// <summary>
    /// Custom conversion delegate. Receives the underlying OpenXml <see cref="Cell"/>
    /// and returns the value to assign to the property. When set, this takes
    /// precedence over the built-in cell parsing.
    /// </summary>
    public Func<Cell, TProperty>? Convert { get; set; }
}
