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
            .Column<int>("Id", _ => _.Heading = "Employee ID")
            .Column<string>("Name", _ => _.Heading = "Full Name")
            .Column<string>("Email", _ => _.Heading = "Email Address")
            .Column<Date?>("HireDate", _ => _.Heading = "Hire Date")
            .Column<int>("Salary", _ => _.Heading = "Annual Salary")
            .Column<bool>("IsActive")
            .Column<EmployeeStatus>("Status");

        reader.Convert(stream);

        var first = sheet.Rows[0];

        #endregion

        Assert.That(first["Id"], Is.EqualTo(1));
        Assert.That(first["Name"], Is.EqualTo("John Doe"));
        Assert.That(first["Email"], Is.EqualTo("john@company.com"));
        Assert.That(first["HireDate"], Is.EqualTo(new Date(2020, 1, 15)));
        Assert.That(first["Salary"], Is.EqualTo(75000));
        Assert.That(first["IsActive"], Is.EqualTo(true));
        Assert.That(first["Status"], Is.EqualTo(EmployeeStatus.FullTime));
    }
}
