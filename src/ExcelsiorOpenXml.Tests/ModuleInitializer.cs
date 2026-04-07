public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DerivePathInfo((_, projectDirectory, _, _) => new(projectDirectory.Replace("Aspose", "OpenXml")));
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifierSettings.InitializePlugins();

    }
}
