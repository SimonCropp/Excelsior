namespace Excelsior;

static class SheetParser
{
    public static Dictionary<string, Dictionary<int, string>> ReadMetadata(WorkbookPart workbookPart)
    {
        var result = new Dictionary<string, Dictionary<int, string>>(StringComparer.Ordinal);
        foreach (var part in workbookPart.GetPartsOfType<CustomXmlPart>())
        {
            using var stream = part.GetStream();
            XDocument doc;
            try
            {
                doc = XDocument.Load(stream);
            }
            catch
            {
                continue;
            }

            var root = doc.Root;
            if (root == null || root.Name.NamespaceName != BookBuilder.MetadataNamespace)
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
                    if (!int.TryParse(column.Attribute("index")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
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
        SharedStringTable? sharedStrings,
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

        var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            return;
        }

        Row? headerRow = null;
        var dataRows = new List<Row>();
        foreach (var row in sheetData.Elements<Row>())
        {
            if (headerRow == null)
            {
                headerRow = row;
            }
            else
            {
                dataRows.Add(row);
            }
        }

        if (headerRow == null)
        {
            return;
        }

        var columnByIndex = ResolveColumnByIndex(
            headerRow,
            sharedStrings,
            metadataColumnMap,
            byName,
            byHeading);

        foreach (var row in dataRows)
        {
            var cellByIndex = new Dictionary<int, Cell>();
            foreach (var cell in row.Elements<Cell>())
            {
                var idx = ParseColumnIndex(cell.CellReference?.Value);
                if (idx >= 0)
                {
                    cellByIndex[idx] = cell;
                }
            }

            var values = new Dictionary<string, object?>(columnByIndex.Count, StringComparer.Ordinal);
            foreach (var (index, column) in columnByIndex)
            {
                cellByIndex.TryGetValue(index, out var cell);
                if (TryConvertCell(resolvedSheetName, row, index, cell, column, sharedStrings, errors, out var value))
                {
                    values[column.Name] = value;
                }
            }

            sheet.Receive(values);
        }
    }

    static Dictionary<int, ColumnReadInfo> ResolveColumnByIndex(
        Row headerRow,
        SharedStringTable? sharedStrings,
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

    static bool TryConvertCell(
        string sheetName,
        Row row,
        int columnIndex,
        Cell? cell,
        ColumnReadInfo column,
        SharedStringTable? sharedStrings,
        List<ReadError> errors,
        out object? value)
    {
        var cellRef = cell?.CellReference?.Value
                      ?? $"{SheetContext.GetColumnLetter(columnIndex)}{row.RowIndex?.Value ?? 0}";

        if (column.Convert != null && cell != null)
        {
            try
            {
                value = column.Convert(cell);
                return true;
            }
            catch (Exception exception)
            {
                errors.Add(new(
                    sheetName,
                    (int)(row.RowIndex?.Value ?? 0),
                    column.Name,
                    cellRef,
                    $"Converter delegate threw: {exception.Message}",
                    exception));
                value = null;
                return false;
            }
        }

        if (CellConverter.TryConvert(cell, column.Type, sharedStrings, out value, out var error))
        {
            return true;
        }

        errors.Add(new(
            sheetName,
            (int)(row.RowIndex?.Value ?? 0),
            column.Name,
            cellRef,
            error,
            null));
        return false;
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
                index = index * 26 + (c - 'A' + 1);
                length++;
            }
            else if (c is >= 'a' and <= 'z')
            {
                index = index * 26 + (c - 'a' + 1);
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
