namespace Excelsior;

public interface IDictionarySheetReader
{
    /// <summary>
    /// Parsed rows. Each entry is a dictionary keyed by column name. Empty until
    /// <c>BookReader.Convert</c>/<c>TryConvert</c> has been called.
    /// </summary>
    IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows { get; }

    /// <summary>
    /// Declares a column. <typeparamref name="TProperty"/> drives the default
    /// cell parsing. Use <see cref="DictionaryColumnReadConfig.Convert"/> to
    /// override the default with a custom delegate.
    /// </summary>
    IDictionarySheetReader Column<TProperty>(
        string name,
        Action<DictionaryColumnReadConfig>? configuration = null);
}
