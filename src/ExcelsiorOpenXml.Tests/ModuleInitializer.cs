public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        DerivePathInfo((_, projectDirectory, _, _) => new(projectDirectory.Replace("Aspose", "OpenXml")));
        VerifierSettings.DontScrubDateTimes();
        VerifierSettings.DontScrubGuids();
        VerifierSettings.InitializePlugins();
        VerifierSettings.RegisterStreamComparer(
            "xlsx",
            (_, _, _) => Task.FromResult(new CompareResult(true)));
        VerifyOpenXmlBook.Initialize();
    }
}
