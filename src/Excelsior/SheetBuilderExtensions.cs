namespace Excelsior;

public static class SheetBuilderExtensions
{
    extension<TModel, TStyle>(ISheetBuilder<TModel, TStyle> builder)
        where TModel : class
    {
        public void HeaderText<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            string value) =>
            builder.Column(property, _ => _.HeaderText = value);

        public void Order<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            int? value) =>
            builder.Column(property, _ => _.Order = value);

        public void ColumnWidth<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            double? value) =>
            builder.Column(property, _ => _.ColumnWidth = value);

        public void HeaderStyle<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            Action<TStyle> value) =>
            builder.Column(property, _ => _.HeaderStyle = value);

        public void DataCellStyle<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            Action<TStyle> value) =>
            builder.Column(property, _ => _.DataCellStyle = value);

        public void ConditionalStyling<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            Action<TStyle, TProperty> value) =>
            builder.Column(property, _ => _.ConditionalStyling = value);

        public void Format<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            string value) =>
            builder.Column(property, _ => _.Format = value);

        public void NullDisplayText<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            string value) =>
            builder.Column(property, _ => _.NullDisplayText = value);

        public void TreatAsHtml<TProperty>(
            Expression<Func<TModel, TProperty>> property) =>
            builder.Column(property, _ => _.TreatAsHtml = true);

        public void Render<TProperty>(
            Expression<Func<TModel, TProperty>> property,
            Func<TProperty, string?> value) =>
            builder.Column(property, _ => _.Render = value);
    }
}