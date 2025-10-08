public static class AsposeLicense
{
    [ModuleInitializer]
    public static void Init()
    {
        var licenseText = Environment.GetEnvironmentVariable("AsposeLicense");
        if (licenseText == null)
        {
            throw new("Expected a `AsposeLicense` environment variable");
        }

        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream);
        writer.Write(licenseText);
        writer.Flush();

        var lic = new Aspose.Cells.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }
}