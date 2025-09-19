namespace Excelsior;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    XLColor? alternateRowColor = null,
    Action<IXLStyle>? headerStyle = null,
    Action<IXLStyle>? globalStyle = null)
{
    List<Action<XLWorkbook>> actions = new();

    public SheetBuilder<T> AddSheet<T>(IEnumerable<T> data, string? name = null)
        where T : class
    {
        name ??= $"Sheet{actions.Count + 1}";

        var converter = new SheetBuilder<T>(
            name,
            data.ToList(),
            useAlternatingRowColors,
            alternateRowColor,
            headerStyle,
            globalStyle);
        actions.Add(_ => converter.AddSheet(_));
        return converter;
    }

    public void ToStream(Stream stream)
    {
        using var book = Build();
        book.SaveAs(stream);
    }

    public XLWorkbook Build()
    {
        var book = new XLWorkbook();
        foreach (var action in actions)
        {
            action(book);
        }

        return book;
    }

    static Dictionary<Type, Func<object, string>> renders = new();

    public static void RenderFor<T>(Func<T, string> func) =>
        renders[typeof(T)] = o => func((T) o);

    internal static bool TryRender(Type memberType, object instance, [NotNullWhen(true)] out string? result)
    {
        foreach (var (key, value) in renders)
        {
            if (key.IsInstanceOfType(instance))
            {
                result = value(instance);
                return true;
            }
        }

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