[TestFixture]
public class BookReaderTests
{
    [Test]
    public async Task RoundTrip_Employees()
    {
        #region BookReaderUsage

        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees());
        await builder.ToStream(stream);
        stream.Position = 0;

        var reader = new BookReader();
        var sheet = reader.AddSheet<Employee>();
        reader.Convert(stream);

        var employees = sheet.Rows;

        #endregion

        await Verify(employees);
    }

    [Test]
    public async Task RoundTrip_NamedSheet()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees(), "Staff");
        await builder.ToStream(stream);
        stream.Position = 0;

        var reader = new BookReader();
        var sheet = reader.AddSheet<Employee>("Staff");
        reader.Convert(stream);

        await Verify(sheet.Rows);
    }
}
