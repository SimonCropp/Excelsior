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
        Action<CellStyle>? tableHeadingStyle,
        MainDocumentPart? mainPart)
    {
        var table = new W.Table();
        table.Append(BuildTableProperties());
        table.Append(BuildGrid(columns.Count));
        table.Append(BuildHeaderRow(columns, tableHeadingStyle));

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
                new W.TopMargin
                {
                    Width = "0",
                    Type = W.TableWidthUnitValues.Dxa
                },
                new W.StartMargin
                {
                    Width = "108",
                    Type = W.TableWidthUnitValues.Dxa
                },
                new W.BottomMargin
                {
                    Width = "0",
                    Type = W.TableWidthUnitValues.Dxa
                },
                new W.EndMargin
                {
                    Width = "108",
                    Type = W.TableWidthUnitValues.Dxa
                }));

    static W.TableGrid BuildGrid(int columnCount)
    {
        var grid = new W.TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new W.GridColumn());
        }

        return grid;
    }

    static W.TableRow BuildHeaderRow(List<ColumnConfig<TModel>> columns, Action<CellStyle>? tableHeadingStyle)
    {
        var row = new W.TableRow();
        foreach (var column in columns)
        {
            row.Append(BuildHeaderCell(column, tableHeadingStyle));
        }

        return row;
    }

    static W.TableCell BuildHeaderCell(ColumnConfig<TModel> column, Action<CellStyle>? tableHeadingStyle)
    {
        var style = ResolveHeadingStyle(column, tableHeadingStyle);

        var run = new W.Run(BuildHeaderRunProperties(style), new W.Text(column.Heading)
        {
            Space = SpaceProcessingModeValues.Preserve
        });

        var paragraph = new W.Paragraph(BuildHeaderParagraphProperties(style), run);
        var cell = new W.TableCell(paragraph);

        var cellProperties = BuildHeaderCellProperties(style);
        if (cellProperties != null)
        {
            cell.PrependChild(cellProperties);
        }

        return cell;
    }

    static CellStyle ResolveHeadingStyle(ColumnConfig<TModel> column, Action<CellStyle>? tableHeadingStyle)
    {
        // Preseed with the renderer's header defaults (bold, centered) so callers can layer on
        // additions (background, font color, size) or opt out (set Font.Bold = false) without
        // having to restate what they didn't want to change.
        var style = new CellStyle
        {
            Alignment =
            {
                Horizontal = HorizontalAlignmentValues.Center
            }
        };
        style.Font.Bold = true;

        tableHeadingStyle?.Invoke(style);
        column.HeadingStyle?.Invoke(style);
        return style;
    }

    static W.RunProperties BuildHeaderRunProperties(CellStyle style)
    {
        var properties = new W.RunProperties();
        if (style.Font.Bold)
        {
            properties.Append(new W.Bold());
        }

        if (style.Font.Underline)
        {
            properties.Append(
                new W.Underline
                {
                    Val = W.UnderlineValues.Single
                });
        }

        if (!string.IsNullOrEmpty(style.Font.Color))
        {
            properties.Append(
                new W.Color
                {
                    Val = style.Font.Color
                });
        }

        if (style.Font.Size is { } size)
        {
            // Word uses half-points for font size.
            var halfPoints = ((int)Math.Round(size * 2)).ToString(CultureInfo.InvariantCulture);
            properties.Append(
                new W.FontSize
                {
                    Val = halfPoints
                });
            properties.Append(
                new W.FontSizeComplexScript
                {
                    Val = halfPoints
                });
        }

        if (!string.IsNullOrEmpty(style.Font.Name))
        {
            properties.Append(
                new W.RunFonts
                {
                    Ascii = style.Font.Name,
                    HighAnsi = style.Font.Name
                });
        }

        return properties;
    }

    static W.ParagraphProperties BuildHeaderParagraphProperties(CellStyle style)
    {
        var horizontal = style.Alignment.Horizontal;
        W.JustificationValues justification;
        if (horizontal == HorizontalAlignmentValues.Left)
        {
            justification = W.JustificationValues.Left;
        }
        else if (horizontal == HorizontalAlignmentValues.Right)
        {
            justification = W.JustificationValues.Right;
        }
        else if (horizontal == HorizontalAlignmentValues.Justify)
        {
            justification = W.JustificationValues.Both;
        }
        else
        {
            // Center / General / Fill all default to centered headers, matching the pre-styling behaviour.
            justification = W.JustificationValues.Center;
        }

        return new(
            new W.Justification
            {
                Val = justification
            });
    }

    static W.TableCellProperties? BuildHeaderCellProperties(CellStyle style)
    {
        var properties = new W.TableCellProperties();
        var hasAny = false;

        if (!string.IsNullOrEmpty(style.BackgroundColor))
        {
            properties.Append(
                new W.Shading
                {
                    Val = W.ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = NormaliseColor(style.BackgroundColor)
                });
            hasAny = true;
        }

        if (style.Alignment.Vertical != VerticalAlignmentValues.Bottom)
        {
            W.TableVerticalAlignmentValues vertical;
            if (style.Alignment.Vertical == VerticalAlignmentValues.Top)
            {
                vertical = W.TableVerticalAlignmentValues.Top;
            }
            else if (style.Alignment.Vertical == VerticalAlignmentValues.Center)
            {
                vertical = W.TableVerticalAlignmentValues.Center;
            }
            else
            {
                vertical = W.TableVerticalAlignmentValues.Bottom;
            }

            properties.Append(
                new W.TableCellVerticalAlignment
                {
                    Val = vertical
                });
            hasAny = true;
        }

        return hasAny ? properties : null;
    }

    static string NormaliseColor(string color) =>
        color.StartsWith('#') ? color[1..] : color;

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
        // Excel formulas (=A2*B2 etc.) have no equivalent in Word tables. Fail loudly so the
        // caller knows the column must be restructured rather than silently dropping the formula.
        if (column.Formula != null)
        {
            throw new($"Column '{column.Heading}' has a Formula, which is not supported in Word tables.");
        }

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
