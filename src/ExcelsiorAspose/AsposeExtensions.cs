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
                cells.SetRowHeight(index, Math.Min(409, Math.Round(height) + 1));
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