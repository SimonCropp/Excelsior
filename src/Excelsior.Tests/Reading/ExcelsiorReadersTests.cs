using DocumentFormat.OpenXml.Spreadsheet;

[TestFixture]
public class ExcelsiorReadersTests
{
    static Cell Cell(string text) =>
        new()
        {
            CellValue = new(text)
        };

    static (Action<int, string> handler, List<(int slot, string message)> errors) ErrorCollector()
    {
        var errors = new List<(int slot, string message)>();
        return ((slot, message) => errors.Add((slot, message)), errors);
    }

    [Test]
    public void ReadString_ReturnsText()
    {
        var result = ExcelsiorReaders.ReadString(Cell("hello"), null);
        Assert.That(result, Is.EqualTo("hello"));
    }

    [Test]
    public void ReadString_TrimsByDefault()
    {
        var result = ExcelsiorReaders.ReadString(Cell("  spaced  "), null);
        Assert.That(result, Is.EqualTo("spaced"));
    }

    [Test]
    public void ReadString_NullCell_ReturnsEmptyString()
    {
        var result = ExcelsiorReaders.ReadString(null, null);
        Assert.That(result, Is.EqualTo(""));
    }

    [Test]
    public void ReadInt_Parses()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadInt(Cell("42"), null, 0, handler);
        Assert.That(result, Is.EqualTo(42));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadInt_EmptyCell_ReportsErrorAtSlot()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadInt(null, null, 7, handler);
        Assert.That(result, Is.EqualTo(0));
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].slot, Is.EqualTo(7));
        Assert.That(errors[0].message, Does.Contain("Int32"));
    }

    [Test]
    public void ReadInt_Unparseable_ReportsErrorAndReturnsDefault()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadInt(Cell("not-a-number"), null, 3, handler);
        Assert.That(result, Is.EqualTo(0));
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].slot, Is.EqualTo(3));
        Assert.That(errors[0].message, Does.Contain("not-a-number"));
    }

    [Test]
    public void ReadIntNullable_EmptyCell_ReturnsNullWithoutError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadIntNullable(null, null, 0, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadIntNullable_Unparseable_ReportsErrorAndReturnsNull()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadIntNullable(Cell("oops"), null, 2, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].slot, Is.EqualTo(2));
    }

    [Test]
    public void ReadDecimal_UsesInvariantCulture()
    {
        var (handler, _) = ErrorCollector();
        var result = ExcelsiorReaders.ReadDecimal(Cell("1234.56"), null, 0, handler);
        Assert.That(result, Is.EqualTo(1234.56m));
    }

    [TestCase("1", true)]
    [TestCase("0", false)]
    [TestCase("true", true)]
    [TestCase("false", false)]
    [TestCase("TRUE", true)]
    [TestCase("FALSE", false)]
    public void ReadBool_AcceptsCommonForms(string raw, bool expected)
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadBool(Cell(raw), null, 0, handler);
        Assert.That(result, Is.EqualTo(expected));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadBool_EmptyCell_ReportsError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadBool(null, null, 0, handler);
        Assert.That(result, Is.False);
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].message, Does.Contain("Boolean"));
    }

    [Test]
    public void ReadBool_Unparseable_ReportsError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadBool(Cell("maybe"), null, 0, handler);
        Assert.That(result, Is.False);
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].message, Does.Contain("maybe"));
    }

    [Test]
    public void ReadBoolNullable_EmptyCell_ReturnsNull()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadBoolNullable(null, null, 0, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadChar_SingleCharacter()
    {
        var (handler, _) = ErrorCollector();
        var result = ExcelsiorReaders.ReadChar(Cell("X"), null, 0, handler);
        Assert.That(result, Is.EqualTo('X'));
    }

    [Test]
    public void ReadChar_MultipleCharacters_ReportsError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadChar(Cell("XY"), null, 0, handler);
        Assert.That(result, Is.EqualTo('\0'));
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].message, Does.Contain("single character"));
    }

    [Test]
    public void ReadCharNullable_EmptyCell_ReturnsNull()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadCharNullable(null, null, 0, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadDateTime_ParsesIsoString()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadDateTime(Cell("2026-05-07T10:30:00"), null, 0, handler);
        Assert.That(result, Is.EqualTo(new DateTime(2026, 5, 7, 10, 30, 0)));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadDateTime_ParsesOaDateNumber()
    {
        var oa = new DateTime(2026, 5, 7).ToOADate().ToString(CultureInfo.InvariantCulture);
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadDateTime(Cell(oa), null, 0, handler);
        Assert.That(result, Is.EqualTo(new DateTime(2026, 5, 7)));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadDate_ParsesOaDateNumber()
    {
        var oa = new DateTime(2026, 5, 7).ToOADate().ToString(CultureInfo.InvariantCulture);
        var (handler, _) = ErrorCollector();
        var result = ExcelsiorReaders.ReadDate(Cell(oa), null, 0, handler);
        Assert.That(result, Is.EqualTo(new Date(2026, 5, 7)));
    }

    [Test]
    public void ReadGuid_Parses()
    {
        var guid = Guid.NewGuid();
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadGuid(Cell(guid.ToString()), null, 0, handler);
        Assert.That(result, Is.EqualTo(guid));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadTimeSpan_ParsesNumericDays()
    {
        var (handler, _) = ErrorCollector();
        var result = ExcelsiorReaders.ReadTimeSpan(Cell("1.5"), null, 0, handler);
        Assert.That(result, Is.EqualTo(TimeSpan.FromDays(1.5)));
    }

    public enum Color
    {
        Red,
        Green,
        Blue
    }

    [Test]
    public void ReadEnum_ByName()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadEnum<Color>(Cell("Green"), null, 0, handler);
        Assert.That(result, Is.EqualTo(Color.Green));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadEnum_CaseInsensitive()
    {
        var (handler, _) = ErrorCollector();
        var result = ExcelsiorReaders.ReadEnum<Color>(Cell("blue"), null, 0, handler);
        Assert.That(result, Is.EqualTo(Color.Blue));
    }

    [Test]
    public void ReadEnum_Unknown_ReportsError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadEnum<Color>(Cell("Yellow"), null, 4, handler);
        Assert.That(result, Is.EqualTo(default(Color)));
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].slot, Is.EqualTo(4));
        Assert.That(errors[0].message, Does.Contain("Yellow"));
    }

    [Test]
    public void ReadEnumNullable_EmptyCell_ReturnsNull()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadEnumNullable<Color>(null, null, 0, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadObject_ForwardsToCellConverter()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadObject(Cell("99"), null, typeof(int), 0, handler);
        Assert.That(result, Is.EqualTo(99));
        Assert.That(errors, Is.Empty);
    }

    [Test]
    public void ReadObject_UnsupportedType_ReportsError()
    {
        var (handler, errors) = ErrorCollector();
        var result = ExcelsiorReaders.ReadObject(Cell("x"), null, typeof(Uri), 5, handler);
        Assert.That(result, Is.Null);
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].slot, Is.EqualTo(5));
    }

    [Test]
    public void SharedStringLookup_ResolvesIndex()
    {
        var cell = new Cell
        {
            DataType = CellValues.SharedString,
            CellValue = new("1")
        };
        var sharedStrings = new[] { "first", "second" };
        var result = ExcelsiorReaders.ReadString(cell, sharedStrings);
        Assert.That(result, Is.EqualTo("second"));
    }
}
