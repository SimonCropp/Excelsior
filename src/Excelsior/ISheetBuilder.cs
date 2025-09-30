namespace Excelsior;

public interface ISheetBuilder<TModel,TStyle>
    where TModel : class
{
    internal void Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<TStyle, TProperty>> configuration);
}