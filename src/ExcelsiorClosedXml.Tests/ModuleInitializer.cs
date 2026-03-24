public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DerivePathInfo((_, projectDirectory, _, _) => new(projectDirectory.Replace("Aspose", "CloseXml")));
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageSharp.Initialize();
        VerifierSettings.InitializePlugins();
    }
}