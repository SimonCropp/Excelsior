static class VerifyOpenXmlBook
{
    public static void Initialize() =>
        VerifierSettings.RegisterFileConverter<OpenXmlBook>(Convert);

    static ConversionResult Convert(OpenXmlBook book, IReadOnlyDictionary<string, object> context)
    {
        var stream = new MemoryStream();
        book.SaveAs(stream);
        stream.Position = 0;
        return new(null, [new("xlsx", stream)]);
    }
}
