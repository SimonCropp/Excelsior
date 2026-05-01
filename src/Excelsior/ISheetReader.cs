namespace Excelsior;

public interface ISheetReader<TModel>
{
    /// <summary>
    /// Parsed rows. Empty until <c>BookReader.Convert</c>/<c>TryConvert</c> has been called.
    /// </summary>
    IReadOnlyList<TModel> Rows { get; }

    ISheetReader<TModel> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnReadConfig<TProperty>> configuration);

    void HeadingText<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value);

    void Order<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    void Include<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        bool value);

    void Exclude<TProperty>(
        Expression<Func<TModel, TProperty>> property);

    /// <summary>
    /// Provide a custom converter for a specific column. The delegate receives
    /// the underlying OpenXml <see cref="Cell"/> and returns the value to assign.
    /// </summary>
    void Convert<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Func<Cell, TProperty> convert);
}
