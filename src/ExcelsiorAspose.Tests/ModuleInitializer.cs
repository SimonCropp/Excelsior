[assembly: NonParallelizable]
public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        ApplyAsposeLicense();

        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageMagick.RegisterComparers(threshold: 0.5);
        VerifierSettings.InitializePlugins();
        VerifierSettings.IgnoreMember("Width");
    }

    static void ApplyAsposeLicense()
    {
        var licenseText = Environment.GetEnvironmentVariable("AsposeLicense");
        if (licenseText == null)
        {
            throw new("Expected a `AsposeLicense` environment variable");
        }

        var stream = new MemoryStream();
        var writer = new StreamWriter(stream);
        writer.Write(licenseText);
        writer.Flush();

        var lic = new Aspose.Cells.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }
}