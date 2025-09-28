namespace Excelsior;

public interface ISheetBuilder<T,TStyle>
    where T : class
{

    internal void Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<TStyle, TProperty>> configuration);
}