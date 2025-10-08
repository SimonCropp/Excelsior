using Application = Microsoft.Office.Interop.Excel.Application;

[TestFixture]
[Apartment(ApartmentState.STA)]
public class RenderExcel
{
    [Test]
    [Explicit]
    public void Run()
    {
        var directory = AttributeReader.GetSolutionDirectory();
        var imageFiles = Directory.EnumerateFiles(directory, "*.png", SearchOption.AllDirectories).ToList();

        foreach (var file in imageFiles)
        {
            if (file.Contains(".verified.") ||
                file.EndsWith("icon.png"))
            {
                continue;
            }

            File.Delete(file);
        }

        var excelFiles = Directory.EnumerateFiles(directory, "*.verified.xlsx", SearchOption.AllDirectories).ToList();
        foreach (var file in excelFiles)
        {
            var name = Path.GetFileName(file);
            if (name.StartsWith("~"))
            {
                continue;
            }

            Convert(file);
        }
    }

    [Test]
    [Explicit]
    public void RunSingle()
    {
        var directory = AttributeReader.GetSolutionDirectory();
        var path = Path.Combine(directory,@"ExcelsiorClosedXml.Tests\Tests.Simple.verified.xlsx");
        Convert(path);
    }

    static void Convert(string excelPath)
    {
        Thread.Sleep(100);
        Application? excel = null;
        Workbook? book = null;

        try
        {
            excel = new()
            {
                Visible = true,
                DisplayAlerts = true,
                ScreenUpdating = true
            };

            // Open workbook
            book = excel.Workbooks.Open(excelPath);

            foreach (Worksheet sheet in book.Sheets)
            {
                RenderSheet(excelPath, sheet);
            }
        }
        catch (Exception exception)
        {
            throw new("Failed for: " + excelPath, exception);
        }
        finally
        {
            try
            {
                if (book != null)
                {
                    Marshal.ReleaseComObject(book);
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    static void RenderSheet(string excelPath, Worksheet sheet)
    {
        try
        {
            var range = sheet.UsedRange;
            sheet.Activate();
            range.Select();
            Clipboard.Clear();
            Thread.Sleep(100);
            range.Copy();
            for (var i = 0; i < 10; i++)
            {
                try
                {
                    range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

                    break;
                }
                catch
                {
                }
                finally
                {
                    Thread.Sleep(500);
                }
            }

            using var image = Clipboard.GetImage()!;
            var imageFile = excelPath
                .Replace(".verified", "")
                .Replace(".xlsx", $"_{sheet.Name}.png");
            image.Save(imageFile, ImageFormat.Png);
        }
        finally
        {
            Marshal.ReleaseComObject(sheet);
        }
    }
}