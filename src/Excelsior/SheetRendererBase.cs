abstract class SheetRendererBase<TModel, TSheet, TStyle, TCell, TBook>(
    IAsyncEnumerable<TModel> data,
    List<Column<TStyle, TModel>> columns)
{
    public List<Column<TStyle, TModel>> Columns => columns;
    protected abstract void SetDateFormat(TStyle style, string format);
    protected abstract void SetNumberFormat(TStyle style, string format);
    protected abstract void SetCellValue(TCell cell, object value);
    protected abstract void SetCellHtml(TCell cell, string value);
    internal abstract Task AddSheet(TBook book, Cancel cancel);
    protected abstract void WriteEnumerable(TCell cell, IEnumerable<string> enumerable);

    protected async Task PopulateData(TSheet sheet, Cancel cancel)
    {
        //Skip heading
        var startRow = 1;

        var itemIndex = 0;
        await foreach (var item in data.WithCancellation(cancel))
        {
            var rowIndex = startRow + itemIndex;

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var value = column.GetValue(item);
                var cell = GetCell(sheet, rowIndex, columnIndex);
                RenderCell(value, column, item, itemIndex, cell);
            }

            itemIndex++;
        }
    }

    protected abstract TCell GetCell(TSheet sheet, int row, int column);

    protected abstract void RenderCell(object? value, Column<TStyle, TModel> column, TModel item, int rowIndex, TCell cell);
    internal void SetCellValue(
        TCell cell,
        TStyle style,
        object? value,
        Column<TStyle, TModel> column,
        TModel item,
        bool trimWhitespace)
    {
        void SetStringOrHtml(string content)
        {
            if (column.IsHtml)
            {
                SetCellHtml(cell, content);
            }
            else
            {
                SetCellValue(cell, content);
            }
        }

        void ThrowIfHtml()
        {
            if (column.IsHtml)
            {
                throw new("TreatAsHtml is not compatible with this type");
            }
        }

        if (value == null)
        {
            if (column.NullDisplay != null)
            {
                SetCellValue(cell, column.NullDisplay);
            }

            return;
        }

        if (column.Render != null)
        {
            var render = column.Render(item, value);
            if (render != null)
            {
                SetStringOrHtml(render);
            }

            return;
        }

        if (value is DateTime dateTime)
        {
            ThrowIfHtml();
            SetCellValue(cell, dateTime);
            SetDateFormat(style, column.Format ?? ValueRenderer.DefaultDateTimeFormat);

            return;
        }

        if (value is Date date)
        {
            ThrowIfHtml();
            SetCellValue(cell, date.ToDateTime(new(0, 0)));
            SetDateFormat(style, column.Format ?? ValueRenderer.DefaultDateFormat);

            return;
        }

        if (value is bool boolean)
        {
            ThrowIfHtml();
            SetCellValue(cell, boolean);
            return;
        }

        if (value is Enum enumValue)
        {
            ThrowIfHtml();
            SetCellValue(cell, enumValue.DisplayName());
            return;
        }

        if (column.IsNumber)
        {
            ThrowIfHtml();
            SetCellValue(cell, Convert.ToDouble(value));
            if (column.Format != null)
            {
                SetNumberFormat(style, column.Format);
            }

            return;
        }

        if (value is IEnumerable<string> enumerable)
        {
            ThrowIfHtml();
            WriteEnumerable(cell, enumerable);
            return;
        }

        var valueAsString = value.ToString();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
        if (valueAsString != null && trimWhitespace)
        {
            valueAsString = valueAsString.Trim();
        }

        if (valueAsString != null)
        {
            SetStringOrHtml(valueAsString);
        }
    }
}