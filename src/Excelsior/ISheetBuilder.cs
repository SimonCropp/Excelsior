namespace Excelsior;

public interface ISheetBuilder<TModel>
{
    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public ISheetBuilder<TModel> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnConfig<TModel, TProperty>> configuration);

    public void HeadingText<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value);

    public void Order<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    public void Width<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    public void MinWidth<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    public void MaxWidth<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    public void HeadingStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<CellStyle> value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<CellStyle, TModel, TProperty> value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<CellStyle, TProperty> value);

    public void Format<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value);

    public void NullDisplay<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value);

    public void IsHtml<TProperty>(
        Expression<Func<TModel, TProperty>> property);

    public void Render<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Func<TModel, TProperty, string?> value);

    public void Render<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Func<TProperty, string?> value);

    /// <summary>
    /// Enable auto-filter for a specific column. Overrides the sheet-level default.
    /// </summary>
    public void Filter<TProperty>(
        Expression<Func<TModel, TProperty>> property);

    /// <summary>
    /// Disable auto-filter for all columns. Individual columns can still opt in via <see cref="Filter{TProperty}"/>
    /// or by setting <see cref="ColumnConfig{TModel,TProperty}.Filter"/> to true.
    /// </summary>
    public void DisableFilter();

    /// <summary>
    /// Include or exclude a specific column from the output.
    /// </summary>
    public void Include<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        bool value);

    /// <summary>
    /// Exclude a specific column from the output.
    /// </summary>
    public void Exclude<TProperty>(
        Expression<Func<TModel, TProperty>> property);
}
