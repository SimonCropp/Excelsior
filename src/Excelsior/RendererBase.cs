using System.Globalization;

abstract class RendererBase<TModel, TSheet, TStyle, TCell, TBook, TColor, TColumn>(
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<TStyle, TModel>> columns,
    int? maxColumnWidth,
    BookBuilderBase<TBook, TSheet, TStyle, TCell, TColor, TColumn> bookBuilder)
{
    protected abstract void SetDateFormat(TStyle style, string format);
    protected abstract void SetStyleColor(TStyle style, TColor color);
    protected abstract void SetNumberFormat(TStyle style, string format);
    protected abstract void SetCellValue(TCell cell, object value);
    protected abstract void SetCellValue(TCell cell, string value);
    protected abstract void SetCellHtml(TCell cell, string value);
    protected abstract TSheet BuildSheet(TBook book);

    protected abstract TColumn GetColumn(TSheet sheet, int index);
    protected abstract void SetColumnWidth(TColumn column, int width);
    protected abstract double GetColumnWidth(TColumn column);

    internal async Task AddSheet(TBook book, Cancel cancel)
    {
        var sheet = BuildSheet(book);
        CreateHeadings(sheet);
        FreezeHeader(sheet);
        await PopulateData(sheet, cancel);
        if (bookBuilder.GlobalStyle != null)
        {
            ApplyGlobalStyling(sheet, bookBuilder.GlobalStyle);
        }

        ApplyFilter(sheet);
        AutoSizeColumns(sheet);
        ResizeRows(sheet);
    }

    protected abstract void FreezeHeader(TSheet sheet);

    void CreateHeadings(TSheet sheet)
    {
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];

            var cell = GetCell(sheet, 0, i);

            SetCellValue(cell, column.Heading);
            var style = GetStyle(cell);
            ApplyHeadingStyling(column, style);
            CommitStyle(cell, style);
        }
    }

    void ApplyHeadingStyling(ColumnConfig<TStyle, TModel> column, TStyle style)
    {
        bookBuilder.HeadingStyle?.Invoke(style);

        column.HeadingStyle?.Invoke(style);
    }

    protected abstract void ApplyGlobalStyling(TSheet sheet, Action<TStyle> bookBuilderGlobalStyle);

    protected abstract void ApplyFilter(TSheet sheet);
    protected abstract void ResizeColumn(TSheet sheet, int index, ColumnConfig<TStyle, TModel> column, int defaultMaxColumnWidth);
    protected abstract void ResizeRows(TSheet sheet);

    void AutoSizeColumns(TSheet sheet)
    {
        var resultMaxColumnWidth = maxColumnWidth ?? bookBuilder.DefaultMaxColumnWidth;
        for (var index = 0; index < columns.Count; index++)
        {
            var column = columns[index];
            ResizeColumn(sheet, index, column, resultMaxColumnWidth);
        }
    }

    async Task PopulateData(TSheet sheet, Cancel cancel)
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
                var style = GetStyle(cell);
                ApplyDefaultStyles(style);
                SetCellValue(cell, style, value, column, item);

                if (bookBuilder.UseAlternatingRowColors &&
                    rowIndex % 2 == 1)
                {
                   SetStyleColor(style, bookBuilder.AlternateRowColor!);
                }

                column.CellStyle?.Invoke(style, item, value);
                CommitStyle(cell, style);
            }

            itemIndex++;
        }
    }

    protected abstract TCell GetCell(TSheet sheet, int row, int column);

    protected abstract void ApplyDefaultStyles(TStyle style);
    protected abstract TStyle GetStyle(TCell cell);
    protected abstract void CommitStyle(TCell cell, TStyle style);

    void SetCellValue(
        TCell cell,
        TStyle style,
        object? value,
        ColumnConfig<TStyle, TModel> column,
        TModel item)
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

        if (value is DateTimeOffset dateTimeOffset)
        {
            ThrowIfHtml();

            SetDateFormat(style, column.Format ?? ValueRenderer.DefaultDateFormat);
            var format = column.Format ?? ValueRenderer.DefaultDateTimeOffsetFormat;
            SetCellValue(cell, dateTimeOffset.ToString(format, CultureInfo.InvariantCulture));

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
            SetCellValue(cell, ListBuilder.Build(enumerable, bookBuilder.TrimWhitespace));
            return;
        }

        var valueAsString = value.ToString();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
        if (valueAsString != null && bookBuilder.TrimWhitespace)
        {
            valueAsString = valueAsString.Trim();
        }

        if (valueAsString != null)
        {
            SetStringOrHtml(valueAsString);
        }
    }
}