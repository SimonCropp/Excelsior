class StyleManager
{
    record FontKey(bool Bold, string? Color, double? Size, string? Name);

    record FillKey(string? BackgroundColor);

    record CellFormatKey(
        uint FontId,
        uint FillId,
        uint? NumberFormatId,
        HorizontalAlignmentValues HAlign,
        VerticalAlignmentValues VAlign,
        bool WrapText);

    List<FontKey> fonts = [new(false, null, null, null)];

    Dictionary<FontKey, uint> fontIndex = new()
    {
        [new(false, null, null, null)] = 0
    };

    List<FillKey> fills = [new(null), new("gray125")];

    Dictionary<FillKey, uint> fillIndex = new()
    {
        [new(null)] = 0
    };

    List<string> customNumberFormats = [];
    Dictionary<string, uint> numberFormatIds = [];
    uint nextNumberFormatId = 164;

    List<CellFormatKey> cellFormats =
        [new(0, 0, null, HorizontalAlignmentValues.General, VerticalAlignmentValues.Bottom, false)];

    Dictionary<CellFormatKey, uint> cellFormatIndex = [];

    internal uint GetOrCreateStyleIndex(CellStyle style)
    {
        var fontId = GetOrCreateFontId(style.Font);
        var fillId = GetOrCreateFillId(style.Fill);
        uint? nfId;
        if (style.NumberFormat == null)
        {
            nfId = null;
        }
        else
        {
            nfId = GetOrCreateNumberFormatId(style.NumberFormat);
        }

        var key = new CellFormatKey(
            fontId,
            fillId,
            nfId,
            style.Alignment.Horizontal,
            style.Alignment.Vertical,
            style.Alignment.WrapText);

        if (cellFormatIndex.TryGetValue(key, out var index))
        {
            return index;
        }

        index = (uint)cellFormats.Count;
        cellFormats.Add(key);
        cellFormatIndex[key] = index;
        return index;
    }

    uint GetOrCreateFontId(CellFont font)
    {
        var key = new FontKey(font.Bold, font.Color, font.Size, font.Name);
        if (fontIndex.TryGetValue(key, out var id))
        {
            return id;
        }

        id = (uint)fonts.Count;
        fonts.Add(key);
        fontIndex[key] = id;
        return id;
    }

    uint GetOrCreateFillId(CellFill fill)
    {
        var key = new FillKey(fill.BackgroundColor);
        if (key.BackgroundColor == null)
        {
            return 0;
        }

        if (fillIndex.TryGetValue(key, out var id))
        {
            return id;
        }

        id = (uint)fills.Count;
        fills.Add(key);
        fillIndex[key] = id;
        return id;
    }

    uint GetOrCreateNumberFormatId(string format)
    {
        if (numberFormatIds.TryGetValue(format, out var id))
        {
            return id;
        }

        id = nextNumberFormatId++;
        numberFormatIds[format] = id;
        customNumberFormats.Add(format);
        return id;
    }

    internal Stylesheet BuildStylesheet()
    {
        var stylesheet = new Stylesheet();

        if (customNumberFormats.Count > 0)
        {
            var nfs = new NumberingFormats
            {
                Count = (uint)customNumberFormats.Count
            };
            uint nfId = 164;
            foreach (var fmt in customNumberFormats)
            {
                nfs.Append(
                    new NumberingFormat
                    {
                        NumberFormatId = nfId++,
                        FormatCode = fmt
                    });
            }

            stylesheet.Append(nfs);
        }

        var fontsEl = new Fonts
        {
            Count = (uint)fonts.Count
        };

        foreach (var fontKey in fonts)
        {
            var font = new Font();
            if (fontKey.Bold)
            {
                font.Append(new Bold());
            }

            font.Append(
                new FontSize
                {
                    Val = fontKey.Size ?? 11
                });

            font.Append(
                new FontName
                {
                    Val = fontKey.Name ?? "Aptos"
                });

            if (fontKey.Color != null)
            {
                font.Append(
                    new Color
                    {
                        Rgb = fontKey.Color
                    });
            }

            fontsEl.Append(font);
        }

        stylesheet.Append(fontsEl);

        var fillsEl = new Fills
        {
            Count = (uint)fills.Count
        };
        fillsEl.Append(
            new Fill(
                new PatternFill
                {
                    PatternType = PatternValues.None
                }));
        fillsEl.Append(
            new Fill(
                new PatternFill
                {
                    PatternType = PatternValues.Gray125
                }));
        for (var i = 2; i < fills.Count; i++)
        {
            fillsEl.Append(new Fill(
                new PatternFill(
                    new ForegroundColor
                    {
                        Rgb = fills[i].BackgroundColor
                    })
                {
                    PatternType = PatternValues.Solid
                }));
        }

        stylesheet.Append(fillsEl);

        var borders = new Borders
        {
            Count = 1
        };
        borders.Append(
            new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()));
        stylesheet.Append(borders);

        var cellFormatsEl = new CellFormats
        {
            Count = (uint)cellFormats.Count
        };
        foreach (var key in cellFormats)
        {
            var cellFormat = new CellFormat
            {
                FontId = key.FontId,
                FillId = key.FillId,
                BorderId = 0,
                ApplyFont = key.FontId > 0,
                ApplyFill = key.FillId > 0
            };

            if (key.NumberFormatId.HasValue)
            {
                cellFormat.NumberFormatId = key.NumberFormatId.Value;
                cellFormat.ApplyNumberFormat = true;
            }

            if (key.HAlign != HorizontalAlignmentValues.General ||
                key.VAlign != VerticalAlignmentValues.Bottom ||
                key.WrapText)
            {
                cellFormat.ApplyAlignment = true;
                cellFormat.Append(new Alignment
                {
                    Horizontal = key.HAlign,
                    Vertical = key.VAlign,
                    WrapText = key.WrapText
                });
            }

            cellFormatsEl.Append(cellFormat);
        }

        stylesheet.Append(cellFormatsEl);

        return stylesheet;
    }
}
