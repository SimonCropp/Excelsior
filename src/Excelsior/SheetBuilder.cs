class SheetBuilder<TModel, TStyle>(Columns<TModel, TStyle> columns) :
    ISheetBuilder<TModel, TStyle>
{
    public ISheetBuilder<TModel, TStyle> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<Column<TStyle, TModel, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    public void HeadingText<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        Column(property, _ => _.Heading = value);

    public void Order<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        int? value) =>
        Column(property, _ => _.Order = value);

    public void Width<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        double? value) =>
        Column(property, _ => _.Width = value);

    public void HeadingStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle> value) =>
        Column(property, _ => _.HeadingStyle = value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TModel, TProperty> value) =>
        Column(property, _ => _.CellStyle = value);

    public void CellStyle<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<TStyle, TProperty> value) =>
        Column(property, _ => _.CellStyle = (style, _, property) => value(style, property));

    public void Format<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        Column(property, _ => _.Format = value);

    public void NullDisplay<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        string value) =>
        Column(property, _ => _.NullDisplay = value);

    public void IsHtml<TProperty>(
        Expression<Func<TModel, TProperty>> property) =>
        Column(property, _ => _.IsHtml = true);

    public void Render<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Func<TModel, TProperty, string?> value) =>
        Column(property, _ => _.Render = value);

    public void Render<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Func<TProperty, string?> value) =>
        Column(property, _ => _.Render = (_, property) => value(property));
}