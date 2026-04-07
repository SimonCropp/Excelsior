class Renderer<TModel>(
    string name,
    IAsyncEnumerable<TModel> data,
    List<ColumnConfig<TModel>> columns,
    int? maxColumnWidth,
    BookBuilder bookBuilder)
{
    internal bool AutoFilter { get; set; } = true;

    List<(CellData Cell, CellStyle Style)> allCells = [];
    Dictionary<int, int> maxContentLength = [];

    internal async Task<SheetData> BuildSheet(Cancel cancel)
    {
        var sheet = new SheetData { Name = name };
        var styleManager = bookBuilder.StyleManager;

        CreateHeadings(sheet, styleManager);
        await PopulateData(sheet, styleManager, cancel);

        if (bookBuilder.GlobalStyle != null)
        {
            ApplyGlobalStyling(styleManager);
        }

        ComputeFilter(sheet);
        ComputeColumnWidths(sheet);
        return sheet;
    }

    void CreateHeadings(SheetData sheet, StyleManager styleManager)
    {
        var row = new List<CellData>();
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var style = new CellStyle();
            style.Font.Bold = true;
            bookBuilder.HeadingStyle?.Invoke(style);
            column.HeadingStyle?.Invoke(style);

            var cellData = MakeInlineStringCell(column.Heading);
            cellData.UpdateStyleIndex(styleManager.GetOrCreateStyleIndex(style));
            row.Add(cellData);
            allCells.Add((cellData, style));
            TrackContentLength(i, column.Heading.Length);
        }

        sheet.Rows.Add(row);
    }

    async Task PopulateData(SheetData sheet, StyleManager styleManager, Cancel cancel)
    {
        var rowIndex = 0;
        await foreach (var item in data.WithCancellation(cancel))
        {
            var row = new List<CellData>();

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var column = columns[columnIndex];
                var value = column.GetValue(item);
                var style = new CellStyle();
                style.Alignment.Horizontal = HorizontalAlignment.Left;
                style.Alignment.Vertical = VerticalAlignment.Top;
                style.Alignment.WrapText = true;

                var cellData = BuildCellValue(value, style, column, item, sheet, rowIndex + 1, columnIndex);

                if (bookBuilder.UseAlternatingRowColors &&
                    rowIndex % 2 == 0)
                {
                    style.BackgroundColor = bookBuilder.AlternateRowColor!;
                }

                column.CellStyle?.Invoke(style, item, value);
                cellData.UpdateStyleIndex(styleManager.GetOrCreateStyleIndex(style));
                row.Add(cellData);
                allCells.Add((cellData, style));
            }

            sheet.Rows.Add(row);
            rowIndex++;
        }
    }

    void ApplyGlobalStyling(StyleManager styleManager)
    {
        foreach (var (cell, style) in allCells)
        {
            bookBuilder.GlobalStyle!(style);
            cell.UpdateStyleIndex(styleManager.GetOrCreateStyleIndex(style));
        }
    }

    void ComputeFilter(SheetData sheet)
    {
        var first = -1;
        var last = -1;
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            if (column.IsEnumerable || !(column.Filter ?? AutoFilter))
            {
                continue;
            }

            if (first == -1)
            {
                first = i;
            }

            last = i;
        }

        if (first != -1 && sheet.Rows.Count > 1)
        {
            sheet.FilterFirstColumn = first;
            sheet.FilterLastColumn = last;
        }
    }

    void ComputeColumnWidths(SheetData sheet)
    {
        var resultMaxColumnWidth = maxColumnWidth ?? bookBuilder.DefaultMaxColumnWidth;

        for (var index = 0; index < columns.Count; index++)
        {
            var columnConfig = columns[index];
            int width;
            if (columnConfig.Width == null)
            {
                maxContentLength.TryGetValue(index, out var contentLen);
                var estimated = contentLen * 1.1 + 2;
                if (estimated < 8)
                {
                    estimated = 8;
                }

                width = (int)Math.Round(estimated);
                width += 1;

                if (columnConfig.IsEnumerable)
                {
                    width += 5;
                }

                if (width > resultMaxColumnWidth)
                {
                    width = resultMaxColumnWidth;
                }
            }
            else
            {
                width = columnConfig.Width.Value;
            }

            sheet.ColumnWidths.Add((index, width));
        }
    }

    void TrackContentLength(int columnIndex, int length)
    {
        if (!maxContentLength.TryGetValue(columnIndex, out var current) || length > current)
        {
            maxContentLength[columnIndex] = length;
        }
    }

    CellData BuildCellValue(
        object? value,
        CellStyle style,
        ColumnConfig<TModel> column,
        TModel item,
        SheetData sheet,
        int rowIndex,
        int columnIndex)
    {
        CellData SetStringOrHtml(string content)
        {
            if (column.IsHtml)
            {
                return MakeHtmlCell(content, columnIndex);
            }

            TrackContentLength(columnIndex, content.Length);
            return MakeInlineStringCell(content);
        }

        if (value == null)
        {
            if (column.NullDisplay != null)
            {
                TrackContentLength(columnIndex, column.NullDisplay.Length);
                return MakeInlineStringCell(column.NullDisplay);
            }

            return new CellData(CellType.Empty, null, 0);
        }

        if (column.TryRender(item, value, out var render))
        {
            return SetStringOrHtml(render);
        }

        if (value is Link link)
        {
            return MakeLinkCell(link, style, sheet, rowIndex, columnIndex);
        }

        if (column.IsEnumerable &&
            value is IEnumerable<Link?> linkEnumerable)
        {
            var links = new List<Link>();
            foreach (var l in linkEnumerable)
            {
                if (l == null)
                {
                    continue;
                }

                links.Add(l);
            }

            if (links.Count > 0)
            {
                return MakeLinkListCell(links, sheet, rowIndex, columnIndex);
            }

            return new CellData(CellType.Empty, null, 0);
        }

        if (column.IsEnumerable &&
            value is IEnumerable enumerable)
        {
            var items = new List<string>();
            foreach (var obj in enumerable)
            {
                if (obj == null)
                {
                    continue;
                }

                var str = column.ItemRender == null ? obj.ToString() : column.ItemRender(obj);
                if (str != null && ValueRenderer.TrimWhitespace)
                {
                    str = str.Trim();
                }

                if (str != null)
                {
                    items.Add(str);
                }
            }

            if (items.Count > 0)
            {
                return MakeListCell(items, columnIndex);
            }

            return new CellData(CellType.Empty, null, 0);
        }

        if (value is DateTime dateTime)
        {
            ThrowIfHtml(column);
            style.NumberFormat = column.Format ?? ValueRenderer.DefaultDateTimeFormat;
            TrackContentLength(columnIndex, 10);
            return new CellData(CellType.Number, dateTime.ToOADate().ToString(CultureInfo.InvariantCulture), 0);
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            ThrowIfHtml(column);
            var format = column.Format ?? ValueRenderer.DefaultDateTimeOffsetFormat;
            style.NumberFormat = format;
            var formatted = dateTimeOffset.ToString(format, CultureInfo.InvariantCulture);
            TrackContentLength(columnIndex, formatted.Length);
            return MakeInlineStringCell(formatted);
        }

        if (value is Date date)
        {
            ThrowIfHtml(column);
            style.NumberFormat = column.Format ?? ValueRenderer.DefaultDateFormat;
            TrackContentLength(columnIndex, 10);
            return new CellData(CellType.Number, date.ToDateTime(new(0, 0)).ToOADate().ToString(CultureInfo.InvariantCulture), 0);
        }

        if (value is bool boolean)
        {
            ThrowIfHtml(column);
            TrackContentLength(columnIndex, 5);
            return new CellData(CellType.Boolean, boolean ? "1" : "0", 0);
        }

        if (column.IsNumber)
        {
            ThrowIfHtml(column);
            if (column.Format != null)
            {
                style.NumberFormat = column.Format;
            }

            var numStr = Convert.ToDouble(value).ToString(CultureInfo.InvariantCulture);
            TrackContentLength(columnIndex, numStr.Length);
            return new CellData(CellType.Number, numStr, 0);
        }

        var valueAsString = value.ToString();
        // ReSharper disable once ConditionIsAlwaysTrueOrFalseAccordingToNullableAPIContract
        if (valueAsString != null && ValueRenderer.TrimWhitespace)
        {
            valueAsString = valueAsString.Trim();
        }

        if (valueAsString != null)
        {
            return SetStringOrHtml(valueAsString);
        }

        return new CellData(CellType.Empty, null, 0);
    }

    static void ThrowIfHtml(ColumnConfig<TModel> column)
    {
        if (column.IsHtml)
        {
            throw new("TreatAsHtml is not compatible with this type");
        }
    }

    static CellData MakeInlineStringCell(string value) =>
        new(CellType.InlineString, InlineStringXml.SimpleText(value), 0);

    CellData MakeHtmlCell(string html, int columnIndex)
    {
        var xml = HtmlToInlineString.Convert(html);
        TrackContentLength(columnIndex, EstimateXmlTextLength(xml));
        return new CellData(CellType.InlineString, xml, 0);
    }

    CellData MakeListCell(IReadOnlyList<string> items, int columnIndex)
    {
        var xml = InlineStringXml.BulletList(items);
        var maxLen = 0;
        foreach (var item in items)
        {
            if (item.Length > maxLen)
            {
                maxLen = item.Length;
            }
        }

        TrackContentLength(columnIndex, maxLen + 2);
        return new CellData(CellType.InlineString, xml, 0);
    }

    CellData MakeLinkCell(Link link, CellStyle style, SheetData sheet, int rowIndex, int columnIndex)
    {
        var display = link.Text ?? link.Url;
        var cellRef = XlsxWriter.GetColumnLetter(columnIndex) + (rowIndex + 1);
        sheet.Hyperlinks.Add(new HyperlinkInfo(cellRef, link.Url));
        style.Font.Color = "0563C1";
        style.Font.Underline = true;
        TrackContentLength(columnIndex, display.Length);
        return new CellData(CellType.InlineString, InlineStringXml.SimpleText(display), 0);
    }

    CellData MakeLinkListCell(List<Link> links, SheetData sheet, int rowIndex, int columnIndex)
    {
        var linkItems = new List<string>(links.Count);
        foreach (var l in links)
        {
            linkItems.Add(l.Text == null ? l.Url : $"{l.Text} ({l.Url})");
        }

        var xml = InlineStringXml.LinkList(linkItems);

        if (links.Count == 1)
        {
            var cellRef = XlsxWriter.GetColumnLetter(columnIndex) + (rowIndex + 1);
            sheet.Hyperlinks.Add(new HyperlinkInfo(cellRef, links[0].Url));
        }

        TrackContentLength(columnIndex, linkItems.Max(_ => _.Length) + 2);
        return new CellData(CellType.InlineString, xml, 0);
    }

    static int EstimateXmlTextLength(string xml)
    {
        // Rough estimate: count characters outside of XML tags
        var length = 0;
        var inTag = false;
        foreach (var c in xml)
        {
            if (c == '<')
            {
                inTag = true;
            }
            else if (c == '>')
            {
                inTag = false;
            }
            else if (!inTag)
            {
                length++;
            }
        }

        return length;
    }
}
