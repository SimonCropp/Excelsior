namespace Excelsior;

public class DictionaryColumnReadConfig
{
    public string? Heading { get; set; }

    /// <summary>
    /// Custom conversion delegate. Receives the underlying OpenXml <see cref="Cell"/>
    /// and returns the value to assign for this column. When set, this takes
    /// precedence over the built-in cell parsing (which infers from the
    /// declared <c>TProperty</c>).
    /// </summary>
    public Func<Cell, object?>? Convert { get; set; }
}
