namespace Excelsior;

public class BookBuilder
{
    public BookBuilder(
        bool useAlternatingRowColors = false,
        string? alternateRowColor = null,
        Action<CellStyle>? headingStyle = null,
        Action<CellStyle>? globalStyle = null,
        int defaultMaxColumnWidth = 50)
    {
        ValueRenderer.SetBookBuilderUsed();
        UseAlternatingRowColors = useAlternatingRowColors;
        AlternateRowColor = alternateRowColor;
        HeadingStyle = headingStyle;
        GlobalStyle = globalStyle;
        DefaultMaxColumnWidth = defaultMaxColumnWidth;
    }

    public bool UseAlternatingRowColors { get; }

    List<Func<SpreadsheetDocument, Cancel, Task>> actions = [];
    public int DefaultMaxColumnWidth { get; }
    public string? AlternateRowColor { get; }
    public Action<CellStyle>? HeadingStyle { get; }
    public Action<CellStyle>? GlobalStyle { get; }

    internal StyleManager StyleManager { get; } = new();

    public ISheetBuilder<TModel> AddSheet<TModel>(
        IEnumerable<TModel> data,
        string? name = null,
        int? defaultMaxColumnWidth = null) =>
        AddSheet(data.ToAsyncEnumerable(), name, defaultMaxColumnWidth);

    public ISheetBuilder<TModel> AddSheet<TModel>(
        IAsyncEnumerable<TModel> data,
        string? name = null,
        int? defaultMaxColumnWidth = null)
    {
        name ??= $"Sheet{actions.Count + 1}";
        var columns = new Columns<TModel>();
        var builder = new SheetBuilder<TModel>(columns);

        actions.Add((book, cancel) =>
        {
            var renderer = new Renderer<TModel>(
                name,
                data,
                columns.OrderedColumns(),
                defaultMaxColumnWidth,
                this)
            {
                AutoFilter = columns.AutoFilter
            };
            return renderer.AddSheet(book, cancel);
        });

        return builder;
    }

    public async Task<SpreadsheetDocument> Build(Cancel cancel = default)
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new(new Sheets());

        foreach (var action in actions)
        {
            cancel.ThrowIfCancellationRequested();
            await action(document, cancel);
        }

        ApplyStylesheet(document);
        return document;
    }

    void ApplyStylesheet(SpreadsheetDocument document)
    {
        var stylesPart = document.WorkbookPart!.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault()
                         ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = StyleManager.BuildStylesheet();
    }

    public async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var document = await Build(cancel);

        if (stream.CanRead)
        {
            document.Clone(stream);
        }
        else
        {
            using var temp = new MemoryStream();
            document.Clone(temp);
            temp.Position = 0;
            await temp.CopyToAsync(stream, cancel);
        }
    }

    public async Task ToFile(string path, Cancel cancel = default)
    {
        await using var stream = File.Create(path);
        await ToStream(stream, cancel);
    }
}
