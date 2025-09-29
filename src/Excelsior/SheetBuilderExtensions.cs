// ConvertToExtensionBlock doesnt work properly in net9 yet
// ReSharper disable ConvertToExtensionBlock
namespace Excelsior;

public static class SheetBuilderExtensions
{
    public static void HeaderText<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.HeaderText = value);

    public static void Order<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        int? value)
        where TModel : class =>
        builder.Column(property, _ => _.Order = value);

    public static void ColumnWidth<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        double? value)
        where TModel : class =>
        builder.Column(property, _ => _.ColumnWidth = value);

    public static void HeaderStyle<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value)
        where TModel : class =>
        builder.Column(property, _ => _.HeaderStyle = value);

    public static void DataCellStyle<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value)
        where TModel : class =>
        builder.Column(property, _ => _.DataCellStyle = value);

    public static void ConditionalStyling<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TProperty> value)
        where TModel : class =>
        builder.Column(property, _ => _.ConditionalStyling = value);

    public static void Format<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.Format = value);

    public static void NullDisplayText<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.NullDisplayText = value);

    public static void TreatAsHtml<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property)
        where TModel : class =>
        builder.Column(property, _ => _.TreatAsHtml = true);

    public static void Render<TModel, TStyle, TProperty>(this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Func<TProperty, string?> value)
        where TModel : class =>
        builder.Column(property, _ => _.Render = value);
}