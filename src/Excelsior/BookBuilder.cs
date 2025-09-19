namespace Excelsior;

public class BookBuilder
{
    bool useAlternatingRowColors;
    XLColor? alternateRowColor;
    Action<IXLStyle>? headerStyle;
    Action<IXLStyle>? globalStyle;

    List<Action<XLWorkbook>> actions = new List<Action<XLWorkbook>>();
    public BookBuilder(
        bool useAlternatingRowColors = false,
        XLColor? alternateRowColor = null,
        Action<IXLStyle>? headerStyle = null,
        Action<IXLStyle>? globalStyle = null)
    {
        this.useAlternatingRowColors = useAlternatingRowColors;
        this.alternateRowColor = alternateRowColor;
        this.headerStyle = headerStyle;
        this.globalStyle = globalStyle;
    }

    public ListToExcelConverter<T> AddSheet<T>(List<T> data)
        where T : class
    {
        var converter = new ListToExcelConverter<T>(data);
        actions.Add(_ => converter.AddSheet(_));
        return converter;
    }

    public void ExportToStream(Stream stream)
    {
        using var workbook = CreateWorkbook();
        workbook.SaveAs(stream);
    }

    public XLWorkbook CreateWorkbook()
    {
        var workbook = new XLWorkbook();
        foreach (var action in actions)
        {
            action(workbook);
        }
        return workbook;
    }
}