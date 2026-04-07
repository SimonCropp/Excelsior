// ReSharper disable UnusedParameter.Local
[TestFixture]
public class Saving
{
    [Test]
    public async Task ToStream()
    {
        var data = SampleData.Employees();

        #region ToStream

        var builder = new BookBuilder();
        builder.AddSheet(data);

        var stream = new MemoryStream();
        await builder.ToStream(stream);

        #endregion

        await Verify(stream, extension: "xlsx");
    }
}

// ReSharper disable UnusedParameter.Local

