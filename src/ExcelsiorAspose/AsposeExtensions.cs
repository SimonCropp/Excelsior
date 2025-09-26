static class AsposeExtensions
{
    extension(Sheet sheet)
    {
        public void AutoSizeRows()
        {
            var cells = sheet.Cells;
            sheet.AutoFitRows();
            //Round row since Aspose AutoFitRows is not deterministic
            for (var index = 0; index < cells.MaxRow; index++)
            {
                var height = cells.GetRowHeight(index);
                cells.SetRowHeight(index, Math.Round(height));
            }
        }

        public void AutoSizeColumns()
        {
            sheet.AutoFitColumns();

            //Round widths since aspose AutoFitColumns is not deterministic
            foreach (var column in sheet.Cells.Columns)
            {
                column.Width = Math.Round(column.Width);
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
}