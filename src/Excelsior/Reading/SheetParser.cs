static class SheetParser
{
    public static Dictionary<string, Dictionary<int, string>> ReadMetadata(WorkbookPart workbookPart)
    {
        var result = new Dictionary<string, Dictionary<int, string>>(StringComparer.Ordinal);
        foreach (var part in workbookPart.GetPartsOfType<CustomXmlPart>())
        {
            using var stream = part.GetStream();
            var doc = XDocument.Load(stream);


            var root = doc.Root;
            if (root == null ||
                root.Name.NamespaceName != BookBuilder.MetadataNamespace)
            {
                continue;
            }

            XNamespace ns = BookBuilder.MetadataNamespace;
            foreach (var sheet in root.Elements(ns + "sheet"))
            {
                var sheetName = sheet.Attribute("name")?.Value;
                if (sheetName == null)
                {
                    continue;
                }

                var map = new Dictionary<int, string>();
                foreach (var column in sheet.Elements(ns + "column"))
                {
                    var indexCalue = column.Attribute("index")?.Value;
                    if (!int.TryParse(indexCalue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
                    {
                        continue;
                    }

                    var name = column.Attribute("property")?.Value;
                    if (name == null)
                    {
                        continue;
                    }

                    map[index] = name;
                }

                result[sheetName] = map;
            }
        }

        return result;
    }

    public static void Parse(
        IReaderSheet sheet,
        string resolvedSheetName,
        WorksheetPart worksheetPart,
        string?[]? sharedStrings,
        Dictionary<int, string>? metadataColumnMap,
        List<ReadError> errors)
    {
        sheet.Reset();

        var columns = sheet.Columns();
        var byName = new Dictionary<string, ColumnReadInfo>(columns.Count, StringComparer.Ordinal);
        var byHeading = new Dictionary<string, ColumnReadInfo>(columns.Count, StringComparer.OrdinalIgnoreCase);
        foreach (var info in columns)
        {
            byName[info.Name] = info;
            byHeading[info.Heading] = info;
        }

        // Streaming reader: avoids materializing the entire SheetData DOM.
        // Each Row is loaded on demand and becomes Gen0 garbage after processing.
        using var reader = OpenXmlReader.Create(worksheetPart);

        var headerSeen = false;
        Dictionary<int, int>? fileIndexToSlot = null;
        Cell?[]? cellsBySlot = null;
        Action<int, string>? onError = null;
        Row? currentRow = null;

        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || !reader.IsStartElement)
            {
                continue;
            }

            var row = (Row)reader.LoadCurrentElement()!;

            if (!headerSeen)
            {
                headerSeen = true;
                var columnByIndex = ResolveColumnByIndex(row, sharedStrings, metadataColumnMap, byName, byHeading);

                var resolvedNames = new HashSet<string>(StringComparer.Ordinal);
                foreach (var info in columnByIndex.Values)
                {
                    resolvedNames.Add(info.Name);
                }

                var beforeCount = errors.Count;
                foreach (var declared in columns)
                {
                    if (resolvedNames.Contains(declared.Name))
                    {
                        continue;
                    }

                    errors.Add(new(
                        resolvedSheetName,
                        RowIndex: 0,
                        declared.Name,
                        CellReference: "",
                        $"Column '{declared.Name}' (heading '{declared.Heading}') was not found in the sheet header row."));
                }

                if (errors.Count > beforeCount)
                {
                    return;
                }

                var slotByName = new Dictionary<string, int>(columns.Count, StringComparer.Ordinal);
                for (var i = 0; i < columns.Count; i++)
                {
                    slotByName[columns[i].Name] = i;
                }

                fileIndexToSlot = new(columnByIndex.Count);
                foreach (var (fileIndex, column) in columnByIndex)
                {
                    fileIndexToSlot[fileIndex] = slotByName[column.Name];
                }

                cellsBySlot = new Cell?[columns.Count];

                // Closure built once; `currentRow` is reassigned per row but the
                // closure reads it lazily so no per-row delegate allocation.
                var slotColumns = columns;
                onError = (slot, message) =>
                {
                    var col = slotColumns[slot];
                    var rowIndex = currentRow?.RowIndex?.Value ?? 0;
                    var cellRef = $"{SheetContext.GetColumnLetter(slot)}{rowIndex}";
                    errors.Add(new(resolvedSheetName, (int)rowIndex, col.Name, cellRef, message));
                };

                continue;
            }

            currentRow = row;
            Array.Clear(cellsBySlot!);

            foreach (var cell in row.Elements<Cell>())
            {
                var index = ParseColumnIndex(cell.CellReference?.Value);
                if (index < 0)
                {
                    continue;
                }

                if (fileIndexToSlot!.TryGetValue(index, out var slot))
                {
                    cellsBySlot![slot] = cell;
                }
            }

            sheet.ReceiveRow(cellsBySlot!, sharedStrings, onError!);
        }
    }

    static Dictionary<int, ColumnReadInfo> ResolveColumnByIndex(
        Row headerRow,
        string?[]? sharedStrings,
        Dictionary<int, string>? metadataColumnMap,
        Dictionary<string, ColumnReadInfo> byName,
        Dictionary<string, ColumnReadInfo> byHeading)
    {
        var result = new Dictionary<int, ColumnReadInfo>();

        foreach (var cell in headerRow.Elements<Cell>())
        {
            var index = ParseColumnIndex(cell.CellReference?.Value);
            if (index < 0)
            {
                continue;
            }

            // Prefer the metadata-XML mapping (round-trip case). Indices in the
            // metadata are 1-based to match the column letter; cell-ref indices
            // here are 0-based, so add one when probing.
            if (metadataColumnMap != null &&
                metadataColumnMap.TryGetValue(index + 1, out var propertyName) &&
                byName.TryGetValue(propertyName, out var byMeta))
            {
                result[index] = byMeta;
                continue;
            }

            var headerText = CellConverter.ReadRaw(cell, sharedStrings);
            if (headerText == null)
            {
                continue;
            }

            if (byHeading.TryGetValue(headerText, out var byHead))
            {
                result[index] = byHead;
                continue;
            }

            // Final fallback: the column name itself, in case the heading was
            // overridden but the writer used DisplayName everywhere.
            if (byName.TryGetValue(headerText, out var byPlainName))
            {
                result[index] = byPlainName;
            }
        }

        return result;
    }

    static int ParseColumnIndex(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return -1;
        }

        var index = 0;
        var length = 0;
        foreach (var c in cellReference)
        {
            if (c is >= 'A' and <= 'Z')
            {
                index = index * 26 + (c - 'A') + 1;
                length++;
            }
            else if (c is >= 'a' and <= 'z')
            {
                index = index * 26 + (c - 'a') + 1;
                length++;
            }
            else
            {
                break;
            }
        }

        return length == 0 ? -1 : index - 1;
    }
}
