[assembly: NonParallelizable]
public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifierSettings.UniqueForTargetFramework();
        VerifierSettings.InitializePlugins();
    }
}