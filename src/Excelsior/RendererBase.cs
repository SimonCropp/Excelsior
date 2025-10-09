abstract class RendererBase<TModel, TSheet, TStyle, TCell, TBook>(
    IAsyncEnumerable<TModel> data,
    List<Column<TStyle, TModel>> columns,
    int defaultMaxColumnWidth)
{
    protected List<Column<TStyle, TModel>> Columns => columns;
    protected abstract void SetDateFormat(TStyle style, string format);
    protected abstract void SetNumberFormat(TStyle style, string format);
    protected abstract void SetCellValue(TCell cell, object value);
    protected abstract void SetCellValue(TCell cell, string value);
    protected abstract void SetCellHtml(TCell cell, string value);
    internal abstract Task AddSheet(TBook book, Cancel cancel);
    protected abstract void ResizeColumn(TSheet sheet, int index, Column<TStyle, TModel> column, int defaultMaxColumnWidth);

    protected void AutoSizeColumns(TSheet sheet)
    {
        for (var index = 0; index < Columns.Count; index++)
        {
            var column = Columns[index];
            ResizeColumn(sheet, index, column, defaultMaxColumnWidth);
        }
    }

    protected async Task PopulateData(TSheet sheet, Cancel cancel)
    {
        var itemIndex = 0;
        await foreach (var item in data.WithCancellation(cancel))
        {
            var rowIndex = itemIndex + 1; // +1 to skip heading;

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

        if (column.TryRender(item, value, out var render))
        {
            SetStringOrHtml(render);

            return;
        }

        if (value is DateTime dateTime)
        {
            ThrowIfHtml();
            SetDateFormat(style, column.Format ?? ValueRenderer.DefaultDateTimeFormat);
            SetCellValue(cell, dateTime);

            return;
        }

        if (value is Date date)
        {
            ThrowIfHtml();
            SetDateFormat(style, column.Format ?? ValueRenderer.DefaultDateFormat);
            SetCellValue(cell, date.ToDateTime(new(0, 0)));

            return;
        }

        if (value is bool boolean)
        {
            ThrowIfHtml();
            SetCellValue(cell, boolean);
            return;
        }

        if (column.IsNumber)
        {
            ThrowIfHtml();
            if (column.Format != null)
            {
                SetNumberFormat(style, column.Format);
            }

            SetCellValue(cell, Convert.ToDouble(value));
            return;
        }

        if (value is IEnumerable<string> enumerable)
        {
            ThrowIfHtml();
            SetCellValue(cell, ListBuilder.Build(enumerable, trimWhitespace));
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