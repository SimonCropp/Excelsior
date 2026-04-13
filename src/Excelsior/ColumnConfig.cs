class ColumnConfig<TModel>
{
    public required string Heading { get; set; }
    public required int? Order { get; set; }
    public required int DeclarationIndex { get; init; }
    public required int? Width { get; set; }
    public required int? MinWidth { get; set; }
    public required int? MaxWidth { get; set; }
    public required Action<CellStyle>? HeadingStyle { get; set; }
    public required Action<CellStyle, TModel, object?>? CellStyle { get; set; }
    public required string? Format { get; set; }
    public required string? NullDisplay { get; set; }
    public required Func<TModel, object, string?>? Render { get; set; }
    public required Func<TModel, FormulaContext<TModel>, string>? Formula { get; set; }
    public required bool IsHtml { get; set; }
    public required bool? Filter { get; set; }
    public required bool Include { get; set; }
    public required bool IsNumber { get; init; }
    public required bool IsEnumerable { get; init; }
    public required Func<object, string>? ItemRender { get; init; }
    public required string Name { get; set; }
    public required Func<TModel, object?> GetValue { get; init; }

    public bool TryRender(TModel item, object value, [NotNullWhen(true)] out string? result)
    {
        if (Render != null)
        {
            result = Render(item, value);
            return result != null;
        }

        result = null;
        return false;
    }
}

public class ColumnConfig<TModel, TProperty>
{
    public string? Heading { get; set; }
    public int? Order { get; set; }
    public int? Width { get; set; }
    public int? MinWidth { get; set; }
    public int? MaxWidth { get; set; }
    public Action<CellStyle>? HeadingStyle { get; set; }
    public Action<CellStyle, TModel, TProperty>? CellStyle { get; set; }
    public string? Format { get; set; }
    public string? NullDisplay { get; set; }
    public Func<TModel, TProperty, string?>? Render { get; set; }

    /// <summary>
    /// An Excel formula for this cell. The callback is invoked per-row and
    /// receives the current model and a <see cref="FormulaContext{TModel}"/>
    /// for resolving cell references. Returning a string like
    /// <c>"=A2*B2"</c> or <c>"A2*B2"</c> sets the cell formula. When set,
    /// this takes precedence over the normal value rendering.
    /// </summary>
    public Func<TModel, FormulaContext<TModel>, string>? Formula { get; set; }

    public bool? IsHtml { get; set; }

    /// <summary>
    /// Enable or disable auto-filter for this column. When null (default), the sheet-level default is used.
    /// </summary>
    public bool? Filter { get; set; }

    /// <summary>
    /// Include or exclude this column from the output. When null (default), the column is included.
    /// Set to false to exclude the column.
    /// </summary>
    public bool? Include { get; set; }
}
