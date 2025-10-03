namespace Excelsior;

public interface ISheetBuilder<TModel, TStyle>
{
    public ISheetBuilder<TModel, TStyle> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<TStyle, TModel, TProperty>> configuration);

    public void HeadingText<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value);

    public void Order<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value);

    public void Width<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        double? value);

    public void HeadingStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TModel, TProperty> value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TProperty> value);

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
}