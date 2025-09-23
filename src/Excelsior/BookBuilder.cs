namespace Excelsior;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    XLColor? alternateRowColor = null,
    Action<IXLStyle>? headerStyle = null,
    Action<IXLStyle>? globalStyle = null,
    bool trimWhitespace = true)
{
    List<SheetBuilder> sheets = [];

    public SheetBuilder<T> AddSheet<T>(IEnumerable<T> data, string? name = null)
        where T : class =>
        AddSheet(data.ToAsyncEnumerable(), name);

    public SheetBuilder<T> AddSheet<T>(IAsyncEnumerable<T> data, string? name = null)
        where T : class
    {
        name ??= $"Sheet{sheets.Count + 1}";

        var converter = new SheetBuilder<T>(
            name,
            data,
            useAlternatingRowColors,
            alternateRowColor,
            headerStyle,
            globalStyle,
            trimWhitespace);
        sheets.Add(converter);
        return converter;
    }

    public async Task ToStream(Stream stream, Cancel cancel = default)
    {
        using var book = await Build(cancel);
        book.SaveAs(stream);
    }

    public async Task<Book> Build(Cancel cancel = default)
    {
        var book = new XLWorkbook();
        foreach (var sheet in sheets)
        {
            cancel.ThrowIfCancellationRequested();
            await sheet.AddSheet(book, cancel);
        }

        return book;
    }

    static Dictionary<Type, Func<object, string>> renders = [];

    public static void RenderFor<T>(Func<T, string> func) =>
        renders[typeof(T)] = _ => func((T) _);

    internal static bool TryRender(Type memberType, object instance, [NotNullWhen(true)] out string? result)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsAssignableTo(memberType))
            {
                result = value(instance);
                return true;
            }
        }

        result = null;
        return false;
    }
}