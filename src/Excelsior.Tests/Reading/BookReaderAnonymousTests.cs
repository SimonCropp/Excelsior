[TestFixture]
public class BookReaderAnonymousTests
{
    [Test]
    public async Task Dictionary_RoundTrip()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet(SampleData.Employees());
        await builder.ToStream(stream);
        stream.Position = 0;

        #region BookReaderDictionary

        var reader = new BookReader();
        var sheet = reader.AddSheet();
        sheet
            .Column<int>("Employee ID")
            .Column<string>("Full Name")
            .Column<string>("Email Address")
            .Column<Date?>("Hire Date")
            .Column<int>("Annual Salary")
            .Column<bool>("IsActive")
            .Column<EmployeeStatus>("Status");

        reader.Convert(stream);

        var first = sheet.Rows[0];

        #endregion

        Assert.That(first["Employee ID"], Is.EqualTo(1));
        Assert.That(first["Full Name"], Is.EqualTo("John Doe"));
        Assert.That(first["Email Address"], Is.EqualTo("john@company.com"));
        Assert.That(first["Hire Date"], Is.EqualTo(new Date(2020, 1, 15)));
        Assert.That(first["Annual Salary"], Is.EqualTo(75000));
        Assert.That(first["IsActive"], Is.EqualTo(true));
        Assert.That(first["Status"], Is.EqualTo(EmployeeStatus.FullTime));
    }
}
