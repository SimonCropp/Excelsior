[assembly: NonParallelizable]
public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifierSettings.UniqueForTargetFramework();
        VerifyImageMagick.RegisterComparers(threshold: 0.05);
        VerifierSettings.InitializePlugins();
    }
}