namespace Excelsior;

/// <summary>
/// Builds a Word table from a sequence of model instances, reusing the same property discovery,
/// ordering, and per-column configuration as <see cref="BookBuilder"/>. The result is a single
/// <c>&lt;w:tbl&gt;</c> element ready to be appended to a Word document body.
/// </summary>
/// <remarks>
/// This builder targets one specific use case: rendering a tabular collection inside an existing
/// Word document part. It is intentionally not concerned with document-level concerns (sections,
/// headers/footers, styles) — the caller owns the host document.
/// </remarks>
public class WordTableBuilder<TModel>(IEnumerable<TModel> data)
{
    readonly Columns<TModel> columns = new();

    /// <summary>
    /// Configure a single column. Mirrors <c>ISheetBuilder&lt;TModel&gt;.Column</c>: any settings
    /// not overridden fall back to <see cref="ColumnAttribute"/> on the model property.
    /// </summary>
    public WordTableBuilder<TModel> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnConfig<TModel, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    /// <summary>
    /// Render the table. When <paramref name="mainPart"/> is supplied, <see cref="Link"/>-typed
    /// values produce real <c>&lt;w:hyperlink&gt;</c> elements with relationships registered on
    /// the host part. When omitted, link cells fall back to their display text only.
    /// </summary>
    public DocumentFormat.OpenXml.Wordprocessing.Table Build(MainDocumentPart? mainPart = null) =>
        WordTableRenderer<TModel>.Build(data, columns.OrderedColumns(), mainPart);
}
