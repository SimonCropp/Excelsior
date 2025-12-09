public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DerivePathInfo((_, projectDirectory, _, _) =>
            new(projectDirectory.Replace("Aspose", "OpenXml")));
        VerifyDiffPlex.Initialize();
        VerifierSettings.InitializePlugins();
        VerifierSettings.ScrubEmptyLines();
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
    }
}
