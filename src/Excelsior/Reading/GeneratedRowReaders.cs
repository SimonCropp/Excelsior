namespace Excelsior;

/// <summary>
/// Registry for source-generated row readers. A row reader takes the raw cells
/// for a row (in slot order matching the model's declared property order), the
/// workbook's shared-string array, and a per-slot error callback. It returns
/// the constructed model with no boxing.
/// </summary>
public static class GeneratedRowReaders
{
    static class Holder<T>
    {
        public static RowReader<T>? Reader;
    }

    public static void Register<T>(RowReader<T> reader) =>
        Holder<T>.Reader = reader;

    public static RowReader<T>? TryGet<T>() =>
        Holder<T>.Reader;
}

/// <summary>
/// Source-generated row reader for type <typeparamref name="T"/>. Errors for
/// individual cells are reported via <paramref name="onError"/> (slot index +
/// human-readable message); the failed property keeps its default value.
/// </summary>
public delegate T RowReader<T>(Cell?[] cellsBySlot, string?[]? sharedStrings, Action<int, string> onError);
