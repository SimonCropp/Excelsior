// ConvertToExtensionBlock doesnt work properly in net9 yet
// ReSharper disable ConvertToExtensionBlock
namespace Excelsior;

public static class SheetBuilderExtensions
{
    public static void HeaderText<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.Header = value);

    public static void Order<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        int? value)
        where TModel : class =>
        builder.Column(property, _ => _.Order = value);

    public static void Width<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        double? value)
        where TModel : class =>
        builder.Column(property, _ => _.Width = value);

    public static void HeaderStyle<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value)
        where TModel : class =>
        builder.Column(property, _ => _.HeaderStyle = value);

    public static void CellStyle<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TModel, TProperty> value)
        where TModel : class =>
        builder.Column(property, _ => _.CellStyle = value);

    public static void CellStyle<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TProperty> value)
        where TModel : class =>
        builder.Column(property, _ => _.CellStyle  = (style, _, property) => value(style, property));

    public static void Format<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.Format = value);

    public static void NullDisplay<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        string value)
        where TModel : class =>
        builder.Column(property, _ => _.NullDisplay = value);

    public static void IsHtml<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property)
        where TModel : class =>
        builder.Column(property, _ => _.IsHtml = true);

    public static void Render<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Func<TModel, TProperty, string?> value)
        where TModel : class =>
        builder.Column(property, _ => _.Render = value);

    public static void Render<TModel, TStyle, TProperty>(
        this ISheetBuilder<TModel, TStyle> builder,
        Expression<Func<TModel, TProperty>> property,
        Func<TProperty, string?> value)
        where TModel : class =>
        builder.Column(property, _ => _.Render = (_, property) => value(property));
}