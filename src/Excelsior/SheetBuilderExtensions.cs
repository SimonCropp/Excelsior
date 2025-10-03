// ConvertToExtensionBlock doesnt work properly in net9 yet
// ReSharper disable ConvertToExtensionBlock
namespace Excelsior;

public static class SheetBuilderExtensions
{
    public static void HeadingText<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        builder.Column(property, _ => _.Heading = value);

    public static void Order<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        int? value) =>
        builder.Column(property, _ => _.Order = value);

    public static void Width<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        double? value) =>
        builder.Column(property, _ => _.Width = value);

    public static void HeadingStyle<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value) =>
        builder.Column(property, _ => _.HeadingStyle = value);

    public static void CellStyle<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TModel, TProperty> value) =>
        builder.Column(property, _ => _.CellStyle = value);

    public static void CellStyle<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TProperty> value) =>
        builder.Column(property, _ => _.CellStyle  = (style, _, property) => value(style, property));

    public static void Format<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        builder.Column(property, _ => _.Format = value);

    public static void NullDisplay<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        builder.Column(property, _ => _.NullDisplay = value);

    public static void IsHtml<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property) =>
        builder.Column(property, _ => _.IsHtml = true);

    public static void Render<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        Func<TModel, TProperty, string?> value) =>
        builder.Column(property, _ => _.Render = value);

    public static void Render<TModel, TStyle, TCell, TProperty>(
        this ISheetBuilder<TModel, TStyle, TCell> builder,
        Expression<Func<TModel, TProperty>> property,
        Func<TProperty, string?> value) =>
        builder.Column(property, _ => _.Render = (_, property) => value(property));
}