namespace Excelsior;

public interface ITemplateSheetBuilder
{
    /// <summary>
    /// Adds a column to the template. Specify <typeparamref name="TProperty"/> to drive the
    /// default cell formatting and validation behavior (e.g. enum types auto-populate
    /// <see cref="TemplateColumnConfig{TProperty}.AllowedValues"/>).
    /// </summary>
    ITemplateSheetBuilder Column<TProperty>(
        string name,
        Action<TemplateColumnConfig<TProperty>>? configuration = null);

    /// <summary>
    /// Disable the default auto-filter on the header row.
    /// </summary>
    void DisableFilter();
}
