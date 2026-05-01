namespace Excelsior;

public class BookBuilder
{
    public BookBuilder(
        bool useAlternatingRowColors = false,
        string? alternateRowColor = null,
        Action<CellStyle>? headingStyle = null,
        Action<CellStyle>? globalStyle = null,
        int? defaultMinColumnWidth = null,
        int defaultMaxColumnWidth = 50,
        int? maxRowHeight = null,
        SheetProtectionOptions? protection = null)
    {
        ValueRenderer.SetBookBuilderUsed();
        UseAlternatingRowColors = useAlternatingRowColors;
        AlternateRowColor = alternateRowColor;
        HeadingStyle = headingStyle;
        GlobalStyle = globalStyle;
        DefaultMinColumnWidth = defaultMinColumnWidth;
        DefaultMaxColumnWidth = defaultMaxColumnWidth;
        MaxRowHeight = maxRowHeight;
        Protection = protection;
    }

    public bool UseAlternatingRowColors { get; }

    List<Func<SpreadsheetDocument, Cancel, Task>> actions = [];
    public int? DefaultMinColumnWidth { get; }
    public int DefaultMaxColumnWidth { get; }
    public int? MaxRowHeight { get; }
    public string? AlternateRowColor { get; }
    public Action<CellStyle>? HeadingStyle { get; }
    public Action<CellStyle>? GlobalStyle { get; }
    public SheetProtectionOptions? Protection { get; }
    internal bool IsProtected => Protection != null;

    internal StyleManager StyleManager { get; } = new();

    public ISheetBuilder<TModel> AddSheet<TModel>(
        IEnumerable<TModel> data,
        string? name = null,
        int? defaultMinColumnWidth = null,
        int? defaultMaxColumnWidth = null,
        int? maxRowHeight = null,
        int templateRowCount = 0,
        bool inferValidationFromTypes = false) =>
        AddSheet(data.ToAsyncEnumerable(), name, defaultMinColumnWidth, defaultMaxColumnWidth, maxRowHeight, templateRowCount, inferValidationFromTypes);

    public ISheetBuilder<TModel> AddSheet<TModel>(
        IAsyncEnumerable<TModel> data,
        string? name = null,
        int? defaultMinColumnWidth = null,
        int? defaultMaxColumnWidth = null,
        int? maxRowHeight = null,
        int templateRowCount = 0,
        bool inferValidationFromTypes = false)
    {
        name ??= $"Sheet{actions.Count + 1}";
        var columns = new Columns<TModel>(inferValidationFromTypes);
        var builder = new SheetBuilder<TModel>(columns);

        actions.Add((book, cancel) =>
        {
            var renderer = new Renderer<TModel>(
                name,
                data,
                columns.OrderedColumns(),
                defaultMinColumnWidth,
                defaultMaxColumnWidth,
                maxRowHeight,
                this,
                templateRowCount)
            {
                AutoFilter = columns.AutoFilter
            };
            return renderer.AddSheet(book, cancel);
        });

        return builder;
    }

    /// <summary>
    /// Adds an empty template sheet with no underlying data binding. Columns are defined explicitly
    /// by name, type, and configuration. Useful when emitting a spreadsheet for users to fill in.
    /// Validation, locked-cell behavior, and conditional formatting all extend down
    /// <paramref name="templateRowCount"/> rows below the header.
    /// </summary>
    public ITemplateSheetBuilder AddTemplateSheet(
        string? name = null,
        int templateRowCount = 1000,
        int? defaultMinColumnWidth = null,
        int? defaultMaxColumnWidth = null,
        int? maxRowHeight = null,
        bool inferValidationFromTypes = true)
    {
        if (templateRowCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(templateRowCount), templateRowCount, "Must be non-negative.");
        }

        name ??= $"Sheet{actions.Count + 1}";
        var builder = new TemplateSheetBuilder(inferValidationFromTypes);

        actions.Add((book, cancel) =>
        {
            var renderer = new Renderer<TemplateRow>(
                name,
                AsyncEnumerable.Empty<TemplateRow>(),
                builder.OrderedColumns(),
                defaultMinColumnWidth,
                defaultMaxColumnWidth,
                maxRowHeight,
                this,
                templateRowCount)
            {
                AutoFilter = builder.AutoFilter
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
        ApplyWorkbookProtection(workbookPart);
        return document;
    }

    void ApplyWorkbookProtection(WorkbookPart workbookPart)
    {
        if (!IsProtected)
        {
            return;
        }

        var sheets = workbookPart.Workbook!.GetFirstChild<Sheets>()!;
        var protection = new WorkbookProtection
        {
            WorkbookPassword = ProtectionPasswordHasher.Hash(Protection!.Password),
            LockStructure = true,
            LockWindows = false
        };
        workbookPart.Workbook.InsertBefore(protection, sheets);
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
