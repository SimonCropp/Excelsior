namespace Excelsior;

public class BookBuilder(
    bool useAlternatingRowColors = false,
    XLColor? alternateRowColor = null,
    Action<IXLStyle>? headerStyle = null,
    Action<IXLStyle>? globalStyle = null)
{
    List<Action<XLWorkbook>> actions = new();

    public SheetBuilder<T> AddSheet<T>(List<T> data, string? name = null)
        where T : class
    {
        name ??= $"Sheet{actions.Count + 1}";

        var converter = new SheetBuilder<T>(
            name,
            data,
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
}