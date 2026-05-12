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

    public class Department
    {
        public required string Name { get; init; }
        public required int HeadCount { get; init; }
    }

    [Test]
    public async Task SheetNameMatchIsCaseInsensitive()
    {
        // Excel sheet names are themselves case-insensitive (you cannot have
        // "Staff" and "staff" in the same workbook), so the reader's sheet
        // lookup must match accordingly.
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees(), "Staff");
        await builder.ToStream(stream);
        stream.Position = 0;

        var reader = new BookReader();
        var sheet = reader.AddSheet<Employee>("staff");
        reader.Convert(stream);

        Assert.That(sheet.Rows, Is.Not.Empty);
    }

    [Test]
    public async Task RoundTrip_MultipleSheets()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees(), "Staff");
        builder.AddSheet<Department>(
            [
                new()
                {
                    Name = "Eng",
                    HeadCount = 12
                },
                new()
                {
                    Name = "Sales",
                    HeadCount = 7
                }
            ],
            "Departments");
        await builder.ToStream(stream);
        stream.Position = 0;

        #region BookReaderMultipleSheets

        var reader = new BookReader();
        var staff = reader.AddSheet<Employee>("Staff");
        var departments = reader.AddSheet<Department>("Departments");
        reader.Convert(stream);

        var employees = staff.Rows;
        var depts = departments.Rows;

        #endregion

        Assert.That(employees, Is.Not.Empty);
        Assert.That(depts.Select(_ => _.Name), Is.EqualTo(["Eng", "Sales"]));
    }
}
