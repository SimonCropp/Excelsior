public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageSharp.Initialize();
        VerifierSettings.InitializePlugins();
        VerifierSettings.IgnoreMember("Width");
    }
}