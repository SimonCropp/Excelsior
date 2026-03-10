namespace ExcelsiorOpenXml;

public class OpenXmlBook : IDisposable
{
    readonly MemoryStream backingStream;

    internal SpreadsheetDocument Document { get; }
    internal StyleManager StyleManager { get; }

    internal OpenXmlBook()
    {
        backingStream = new();
        Document = SpreadsheetDocument.Create(backingStream, SpreadsheetDocumentType.Workbook);
        var workbookPart = Document.AddWorkbookPart();
        workbookPart.Workbook = new(new Sheets());
        StyleManager = new();
    }

    public void ApplyStylesheet()
    {
        var stylesPart = Document.WorkbookPart!.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault()
                         ?? Document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = StyleManager.BuildStylesheet();
    }

    public void SaveAs(Stream stream)
    {
        ApplyStylesheet();
        if (stream.CanRead)
        {
            Document.Clone(stream);
        }
        else
        {
            using var temp = new MemoryStream();
            Document.Clone(temp);
            temp.Position = 0;
            temp.CopyTo(stream);
        }
    }

    public void Dispose()
    {
        Document.Dispose();
        backingStream.Dispose();
    }
}
