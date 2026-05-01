namespace Excelsior;

public class BookReader
{
    List<IReaderSheet> sheets = [];

    /// <summary>
    /// Register a strong-typed sheet. Properties are auto-discovered from
    /// <typeparamref name="TModel"/>; per-column overrides are applied via the
    /// returned fluent API.
    /// </summary>
    /// <param name="name">Sheet name to read. When null, sheets are matched
    /// positionally by registration order.</param>
    public ISheetReader<TModel> AddSheet<TModel>(string? name = null)
    {
        var reader = new SheetReader<TModel>(name);
        sheets.Add(reader);
        return reader;
    }

    /// <summary>
    /// Register a sheet whose rows are parsed into <see cref="IReadOnlyDictionary{TKey,TValue}"/>
    /// instances keyed by column name. Columns must be declared explicitly via
    /// <see cref="IDictionarySheetReader.Column{TProperty}"/>.
    /// </summary>
    /// <param name="name">Sheet name to read. When null, sheets are matched
    /// positionally by registration order.</param>
    public IDictionarySheetReader AddSheet(string? name = null)
    {
        var reader = new DictionarySheetReader(name);
        sheets.Add(reader);
        return reader;
    }

    /// <summary>
    /// Reads the workbook and populates each registered sheet's <c>Rows</c>.
    /// Throws <see cref="ReadException"/> with the full error collection if any
    /// cell fails to convert.
    /// </summary>
    public void Convert(Stream stream)
    {
        var errors = ConvertCore(stream);
        if (errors.Count > 0)
        {
            throw new ReadException(errors);
        }
    }

    public void Convert(string path)
    {
        using var stream = File.OpenRead(path);
        Convert(stream);
    }

    /// <summary>
    /// Reads the workbook and populates each registered sheet's <c>Rows</c>.
    /// Never throws on data errors. The returned <see cref="ReadResult"/> is
    /// implicitly convertible to <see cref="bool"/> (success) and to a
    /// collection of <see cref="ReadError"/>.
    /// </summary>
    public ReadResult TryConvert(Stream stream) =>
        new(ConvertCore(stream));

    public ReadResult TryConvert(string path)
    {
        using var stream = File.OpenRead(path);
        return TryConvert(stream);
    }

    List<ReadError> ConvertCore(Stream stream)
    {
        var errors = new List<ReadError>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var metadata = SheetParser.ReadMetadata(workbookPart);

        var workbookSheets = workbookPart.Workbook?
            .GetFirstChild<Sheets>()?
            .Elements<Sheet>()
            .ToList() ?? [];

        for (var i = 0; i < sheets.Count; i++)
        {
            var sheet = sheets[i];
            var match = ResolveSheet(sheet, i, workbookSheets);
            if (match == null)
            {
                errors.Add(new(
                    sheet.Name ?? $"#{i}",
                    0,
                    "",
                    "",
                    sheet.Name == null
                        ? $"Workbook contains fewer than {i + 1} sheets."
                        : $"Workbook does not contain a sheet named '{sheet.Name}'."));
                sheet.Reset();
                continue;
            }

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(match.Id!.Value!);
            var resolvedName = match.Name?.Value ?? sheet.Name ?? $"#{i}";
            metadata.TryGetValue(resolvedName, out var columnMap);
            SheetParser.Parse(sheet, resolvedName, worksheetPart, sharedStrings, columnMap, errors);
        }

        return errors;
    }

    static Sheet? ResolveSheet(IReaderSheet sheet, int index, List<Sheet> workbookSheets)
    {
        if (sheet.Name != null)
        {
            return workbookSheets.FirstOrDefault(_ => _.Name?.Value == sheet.Name);
        }

        return index < workbookSheets.Count ? workbookSheets[index] : null;
    }
}
