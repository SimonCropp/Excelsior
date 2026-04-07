using System.IO.Compression;

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

    List<Func<Cancel, Task<SheetData>>> actions = [];
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

        actions.Add(cancel =>
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
            return renderer.BuildSheet(cancel);
        });

        return builder;
    }

    public async Task<MemoryStream> Build(Cancel cancel = default)
    {
        var sheets = new List<SheetData>();
        foreach (var action in actions)
        {
            cancel.ThrowIfCancellationRequested();
            sheets.Add(await action(cancel));
        }

        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            XlsxWriter.Write(archive, sheets, StyleManager);
        }

        stream.Position = 0;
        return stream;
    }

    public async Task ToStream(Stream stream, Cancel cancel = default)
    {
        var sheets = new List<SheetData>();
        foreach (var action in actions)
        {
            cancel.ThrowIfCancellationRequested();
            sheets.Add(await action(cancel));
        }

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        XlsxWriter.Write(archive, sheets, StyleManager);
    }

    public async Task ToFile(string path, Cancel cancel = default)
    {
        await using var stream = File.Create(path);
        await ToStream(stream, cancel);
    }
}
