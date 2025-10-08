public static class ApplyLicense
{
    [ModuleInitializer]
    public static void Initialize()
    {
        var license = Environment.GetEnvironmentVariable("SyncfusionLicense");
        if (license == null)
        {
            throw new("Expected a `SyncfusionLicense` environment variable");
        }

        SyncfusionLicenseProvider.RegisterLicense(license);
    }
}