public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifyImageMagick.RegisterComparers(threshold: 0.8);
        VerifierSettings.InitializePlugins();
        VerifierSettings.IgnoreMember("Width");
    }
}