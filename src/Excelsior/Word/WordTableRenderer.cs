namespace Excelsior;

using W = DocumentFormat.OpenXml.Wordprocessing;

/// <summary>
/// Internal worker that turns a sequence of model rows + column configs into a Word
/// <c>&lt;w:tbl&gt;</c> element.
/// </summary>
static class WordTableRenderer<TModel>
{
    public static W.Table Build(
        IEnumerable<TModel> data,
        List<ColumnConfig<TModel>> columns,
        MainDocumentPart? mainPart)
    {
        var table = new W.Table();
        table.Append(BuildTableProperties());
        table.Append(BuildGrid(columns.Count));
        table.Append(BuildHeaderRow(columns));

        foreach (var item in data)
        {
            table.Append(BuildDataRow(columns, item, mainPart));
        }

        return table;
    }

    static W.TableProperties BuildTableProperties() =>
        new(
            new W.TableBorders(
                new W.TopBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                },
                new W.BottomBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                },
                new W.LeftBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                },
                new W.RightBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                },
                new W.InsideHorizontalBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                },
                new W.InsideVerticalBorder
                {
                    Val = W.BorderValues.Single,
                    Size = 4
                }),
            new W.TableCellMarginDefault(
                new W.TopMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
                new W.StartMargin { Width = "108", Type = W.TableWidthUnitValues.Dxa },
                new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
                new W.EndMargin { Width = "108", Type = W.TableWidthUnitValues.Dxa }));

    static W.TableGrid BuildGrid(int columnCount)
    {
        var grid = new W.TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new W.GridColumn());
        }

        return grid;
    }

    static W.TableRow BuildHeaderRow(List<ColumnConfig<TModel>> columns)
    {
        var row = new W.TableRow();
        foreach (var column in columns)
        {
            var paragraph = new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Justification
                    {
                        Val = W.JustificationValues.Center
                    }),
                new W.Run(
                    new W.RunProperties(new W.Bold()),
                    new W.Text(column.Heading)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
            row.Append(new W.TableCell(paragraph));
        }

        return row;
    }

    static W.TableRow BuildDataRow(List<ColumnConfig<TModel>> columns, TModel item, MainDocumentPart? mainPart)
    {
        var row = new W.TableRow();
        foreach (var column in columns)
        {
            var value = column.GetValue(item);
            row.Append(new W.TableCell(BuildCellParagraph(column, item, value, mainPart)));
        }

        return row;
    }

    static W.Paragraph BuildCellParagraph(
        ColumnConfig<TModel> column,
        TModel item,
        object? value,
        MainDocumentPart? mainPart)
    {
        // A live MainDocumentPart is required to register the hyperlink relationship; without one,
        // fall through to the plain-text path which renders link.Text ?? link.Url.
        if (value is Link link && mainPart != null)
        {
            return BuildHyperlinkParagraph(link, mainPart);
        }

        var text = ToText(column, item, value);
        return new(
            new W.Run(
                new W.Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
    }

    static W.Paragraph BuildHyperlinkParagraph(Link link, MainDocumentPart mainPart)
    {
        var display = link.Text ?? link.Url;
        var relId = mainPart.AddHyperlinkRelationship(new(link.Url, UriKind.RelativeOrAbsolute), true).Id;

        // Color and underline mirror Excelsior's spreadsheet hyperlink styling (#0563C1, underline)
        // and avoid depending on a "Hyperlink" character style that may not exist in the host doc.
        var run = new W.Run(
            new W.RunProperties(
                new W.Color
                {
                    Val = "0563C1"
                },
                new W.Underline
                {
                    Val = W.UnderlineValues.Single
                }),
            new W.Text(display)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

        return new(
            new W.Hyperlink(run)
            {
                Id = relId
            });
    }

    static string ToText(ColumnConfig<TModel> column, TModel item, object? value)
    {
        if (value == null)
        {
            return column.NullDisplay ?? string.Empty;
        }

        if (column.TryRender(item, value, out var rendered))
        {
            return rendered;
        }

        // Hyperlink fallback when the renderer wasn't given a MainDocumentPart: emit the display
        // text only. The relationship cannot be registered without a host document part.
        if (value is Link link)
        {
            return link.Text ?? link.Url;
        }

        if (column.IsEnumerable && value is IEnumerable enumerable and not string)
        {
            var items = new List<string>();
            foreach (var entry in enumerable)
            {
                if (entry == null)
                {
                    continue;
                }

                var entryText = column.ItemRender == null ? entry.ToString() : column.ItemRender(entry);
                if (entryText != null && ValueRenderer.TrimWhitespace)
                {
                    entryText = entryText.Trim();
                }

                if (entryText != null)
                {
                    items.Add(entryText);
                }
            }

            return string.Join(", ", items);
        }

        if (value is DateTime dateTime)
        {
            return dateTime.ToString(column.Format ?? ValueRenderer.DefaultDateTimeFormat, ValueRenderer.Culture);
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            return dateTimeOffset.ToString(column.Format ?? ValueRenderer.DefaultDateTimeOffsetFormat, ValueRenderer.Culture);
        }

        if (value is Date date)
        {
            return date.ToDateTime(new(0, 0)).ToString(column.Format ?? ValueRenderer.DefaultDateFormat, ValueRenderer.Culture);
        }

        if (column.Format != null && value is IFormattable formattable)
        {
            return formattable.ToString(column.Format, ValueRenderer.Culture);
        }

        var asString = value.ToString() ?? string.Empty;
        if (ValueRenderer.TrimWhitespace)
        {
            asString = asString.Trim();
        }

        return asString;
    }
}
