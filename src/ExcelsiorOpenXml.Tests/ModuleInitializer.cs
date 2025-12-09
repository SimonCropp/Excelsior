using System.Runtime.CompilerServices;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifyDiffPlex.Initialize();
        VerifierSettings.InitializePlugins();
        VerifierSettings.ScrubEmptyLines();
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        Environment.SetEnvironmentVariable("DiffEngine_Disabled", "true");
    }
}
