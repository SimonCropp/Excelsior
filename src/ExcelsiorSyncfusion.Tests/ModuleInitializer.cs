public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DerivePathInfo((_, projectDirectory, _, _) =>
            new(projectDirectory.Replace("Aspose", "Syncfusion")));
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageMagick.RegisterComparers(threshold: 0.05);
        VerifierSettings.InitializePlugins();
    }
}