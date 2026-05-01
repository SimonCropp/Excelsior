[TestFixture]
public class BookReaderErrorTests
{
    public class StringSource
    {
        public required string Number { get; init; }
    }

    public class IntTarget
    {
        public int Number { get; set; }
    }

    static async Task<MemoryStream> WriteStringNumber()
    {
        var stream = new MemoryStream();
        var builder = new BookBuilder();
        builder.AddSheet<StringSource>(
        [
            new() { Number = "42" },
            new() { Number = "not-a-number" },
            new() { Number = "100" }
        ]);
        await builder.ToStream(stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task Convert_Throws_ReadException_With_Errors()
    {
        var stream = await WriteStringNumber();

        var reader = new BookReader();
        reader.AddSheet<IntTarget>();

        var exception = Assert.Throws<ReadException>(() => reader.Convert(stream))!;
        var errors = exception.Errors;
        Assert.That(errors, Has.Count.EqualTo(1));
        Assert.That(errors[0].ColumnName, Is.EqualTo("Number"));
        Assert.That(errors[0].Message, Does.Contain("not-a-number"));
    }

    [Test]
    public async Task TryConvert_Returns_Errors_Without_Throwing()
    {
        #region BookReaderTryConvert

        var stream = await WriteStringNumber();
        var reader = new BookReader();
        var sheet = reader.AddSheet<IntTarget>();

        var result = reader.TryConvert(stream);
        if (!result)
        {
            foreach (var error in result.Errors)
            {
                Console.WriteLine(error);
            }
        }

        #endregion

        Assert.That(result.Succeeded, Is.False);
        Assert.That((bool)result, Is.False);
        var errors = (ReadError[])result;
        Assert.That(errors, Has.Length.EqualTo(1));
        Assert.That(errors[0].ColumnName, Is.EqualTo("Number"));
        Assert.That(sheet.Rows, Has.Count.EqualTo(3));
        Assert.That(sheet.Rows[0].Number, Is.EqualTo(42));
        Assert.That(sheet.Rows[1].Number, Is.EqualTo(0));
        Assert.That(sheet.Rows[2].Number, Is.EqualTo(100));
    }

    [Test]
    public async Task Exception_Errors_Match_TryConvert_Errors()
    {
        var stream1 = await WriteStringNumber();
        var stream2 = await WriteStringNumber();

        var reader1 = new BookReader();
        reader1.AddSheet<IntTarget>();
        var tryResult = reader1.TryConvert(stream1);

        var reader2 = new BookReader();
        reader2.AddSheet<IntTarget>();
        var exception = Assert.Throws<ReadException>(() => reader2.Convert(stream2))!;

        Assert.That(exception.Errors.Select(_ => _.Message), Is.EqualTo(tryResult.Errors.Select(_ => _.Message)));
    }
}
