static class AsposeExtensions
{
    extension(Sheet sheet)
    {
        public void AutoSizeRows()
        {
            var cells = sheet.Cells;
            sheet.AutoFitRows();
            //Round row since Aspose AutoFitRows is not deterministic.
            //Tall rows (multi-line) have larger variance, so use coarser 5-point rounding.
            //MaxRow is expensive so cache
            var maxRow = cells.MaxRow;
            for (var index = 0; index <= maxRow; index++)
            {
                var height = cells.GetRowHeight(index);
                var rounded = height < 40
                    ? Math.Round(height) + 1
                    : Math.Round(height / 5.0) * 5 + 1;
                cells.SetRowHeight(index, Math.Min(409, rounded));
            }
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