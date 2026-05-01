namespace Excelsior;

public interface IDictionarySheetReader
{
    /// <summary>
    /// Parsed rows. Each entry is a dictionary keyed by column name. Empty until
    /// <c>BookReader.Convert</c>/<c>TryConvert</c> has been called.
    /// </summary>
    IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows { get; }

    /// <summary>
    /// Declares a column. <paramref name="name"/> is matched against the file's
    /// header row (case-insensitively) and is also the key under which the value
    /// appears in each row dictionary. For round-tripped files written by
    /// <c>BookBuilder</c>, <paramref name="name"/> can alternatively be the
    /// underlying property name and is resolved via the workbook's metadata.
    /// <typeparamref name="TProperty"/> drives the default cell parsing; use
    /// <see cref="DictionaryColumnReadConfig.Convert"/> to override it.
    /// </summary>
    IDictionarySheetReader Column<TProperty>(
        string name,
        Action<DictionaryColumnReadConfig>? configuration = null);
}
