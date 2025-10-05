static class AsposeExtensions
{
    extension(Sheet sheet)
    {
        public void AutoSizeRows()
        {
            var cells = sheet.Cells;
            sheet.AutoFitRows();
            //Round row since Aspose AutoFitRows is not deterministic
            //MaxRow is expensive so cache
            var maxRow = cells.MaxRow;
            for (var index = 0; index <= maxRow; index++)
            {
                var height = cells.GetRowHeight(index);
                cells.SetRowHeight(index, Math.Round(height) + 1);
            }
        }

        public void AutoFilterAll()
        {
            var cells = sheet.Cells;
            var row = CellsHelper.RowIndexToName(cells.MaxRow);
            var column = CellsHelper.ColumnIndexToName(cells.MaxColumn);
            sheet.AutoFilter.Range = $"A1:{column}{row}";
        }
    }

    extension(Cell cell)
    {
        public void SafeSetHtml(string? value)
        {
            if (value == null)
            {
                cell.PutValue("");
                return;
            }

            try
            {
                cell.HtmlString = value;
            }
            catch
            {
                cell.Value = value;
            }
        }
    }
}