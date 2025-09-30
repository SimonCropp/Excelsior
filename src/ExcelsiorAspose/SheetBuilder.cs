namespace ExcelsiorAspose;

public class SheetBuilder<T>(
    string name,
    IAsyncEnumerable<T> data,
    bool useAlternatingRowColors,
    Color? alternateRowColor,
    Action<Style>? headerStyle,
    Action<Style>? globalStyle,
    bool trimWhitespace) :
    ISheetBuilder<T, Style>
    where T : class
{
    int rowIndex;
    Columns<T, Style> columns = new();

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public SheetBuilder<T> Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<Style, TProperty>> configuration)
    {
        columns.Add(property, configuration);
        return this;
    }

    void ISheetBuilder<T, Style>.Column<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<Column<Style, TProperty>> configuration) =>
        Column(property, configuration);

    internal async Task AddSheet(Book book, Cancel cancel)
    {
        var sheet = book.Worksheets.Add(name);

        var orderedColumns = columns.OrderedColumns();
        CreateHeaders(sheet, orderedColumns);

        await PopulateData(sheet, orderedColumns, cancel);

        ApplyGlobalStyling(sheet);
        sheet.AutoFilterAll();
        AutoSizeColumns(sheet, orderedColumns);
        sheet.AutoSizeRows();
    }

    void CreateHeaders(Sheet sheet, List<Column<Style>> orderedColumns)
    {
        for (var i = 0; i < orderedColumns.Count; i++)
        {
            var column = orderedColumns[i];
            var cell = sheet.Cells[0, i];

            cell.Value = column.HeaderText;

            ApplyHeaderStyling(cell, column);
        }

        sheet.FreezePanes(1, 0, 1, 0);
    }

    async Task PopulateData(Sheet sheet, List<Column<Style>> orderedColumns, Cancel cancel)
    {
        //Skip header
        var startRow = 1;

        await foreach (var item in data.WithCancellation(cancel))
        {
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < orderedColumns.Count; colIndex++)
            {
                var column = orderedColumns[colIndex];

                var cell = sheet.Cells[xlRow, colIndex];

                var style = cell.GetStyle();
                style.VerticalAlignment = TextAlignmentType.Top;
                style.HorizontalAlignment = TextAlignmentType.Left;
                style.IsTextWrapped = true;
                var value = column.GetValue(item);
                SetCellValue(cell, value, style, column);
                ApplyCellStyle(rowIndex, value, style, column);
                cell.SetStyle(style);
            }

            rowIndex++;
        }
    }

    void SetCellValue(Cell cell, object? value, Style style, Column<Style> column)
    {
        if (value == null)
        {
            cell.Value = column.NullDisplayText;
            return;
        }

        if (column.Render != null)
        {
            SetStringOrHtml(column.Render(value));
            return;
        }

        if (value is DateTime dateTime)
        {
            ThrowIfHtml();
            cell.Value = dateTime;
            if (column.Format != null)
            {
                style.Custom = column.Format;
            }

            return;
        }

        if (value is bool boolean)
        {
            ThrowIfHtml();
            cell.Value = boolean.ToString();
            return;
        }

        if (value is Enum enumValue)
        {
            ThrowIfHtml();
            cell.Value = enumValue.DisplayName();
            return;
        }

        if (column.IsNumber)
        {
            ThrowIfHtml();
            cell.Value = Convert.ToDouble(value);
            if (column.Format != null)
            {
                style.Custom = column.Format;
            }

            return;
        }

        if (value is IEnumerable<string> enumerable)
        {
            ThrowIfHtml();
            WriteEnumerable(cell, enumerable);
            return;
        }

        SetStringOrHtml(GetTrimmedValue(value));

        void ThrowIfHtml()
        {
            if (column.TreatAsHtml)
            {
                throw new("TreatAsHtml is not compatible with this type");
            }
        }

        void SetStringOrHtml(string? rendered)
        {
            if (column.TreatAsHtml)
            {
                cell.SafeSetHtml(rendered);
            }
            else
            {
                cell.Value = rendered;
            }
        }
    }

    string? GetTrimmedValue(object value)
    {
        var result = value.ToString();
        if (result != null && trimWhitespace)
        {
            return result.Trim();
        }

        return result;
    }

    void WriteEnumerable(Cell cell, IEnumerable<string> enumerable)
    {
        var list = enumerable.ToList();
        var builder = new StringBuilder(
            """
            <ul>

            """);
        for (var index = 0; index < list.Count; index++)
        {
            var item = list[index];
            builder.Append("<li>");

            if (trimWhitespace)
            {
                item = item.Trim();
            }

            item = WebUtility.HtmlEncode(item);

            // works around a bug where aspose indents only the first item
            if (index != 0)
            {
                item = $"&nbsp;{item}";
            }

            builder.Append(item);

            builder.AppendLine("</li>");
        }

        builder.Append("</ul>");

        cell.SafeSetHtml(builder.ToString());
    }

    void ApplyHeaderStyling(Cell cell, Column<Style> column)
    {
        var style = cell.GetStyle();
        headerStyle?.Invoke(style);

        column.HeaderStyle?.Invoke(style);

        cell.SetStyle(style);
    }

    void ApplyCellStyle(int index, object? value, Style style, Column<Style> column)
    {
        // Apply alternating row colors
        if (useAlternatingRowColors &&
            index % 2 == 1)
        {
            style.BackgroundColor = alternateRowColor!.Value;
        }


        column.CellStyle?.Invoke(style, value);
    }

    void ApplyGlobalStyling(Sheet sheet)
    {
        if (globalStyle == null)
        {
            return;
        }

        var style = sheet.Workbook.CreateStyle();
        globalStyle(style);
        var flag = new StyleFlag
        {
            FontName = true,
            FontSize = true,
            FontColor = true,
            CellShading = true
        };
        sheet.Cells.ApplyStyle(style, flag);
    }

    static void AutoSizeColumns(Sheet sheet, List<Column<Style>> orderedColumns)
    {
        sheet.AutoSizeColumns();

        for (var index = 0; index < orderedColumns.Count; index++)
        {
            var column = orderedColumns[index];
            if (column.ColumnWidth != null)
            {
                sheet.Cells.Columns[index].Width = column.ColumnWidth.Value;
            }
        }
    }
}