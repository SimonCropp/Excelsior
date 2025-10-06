class DisposableBook(ExcelEngine engine, Book book) :
    IDisposableBook
{
    public void Dispose() =>
        engine.Dispose();

    public IApplication Application => book.Application;

    public object Parent => book.Parent;

    public IDataSort CreateDataSorter() =>
        book.CreateDataSorter();

    public void Activate() =>
        book.Activate();

    public IFont AddFont(IFont fontToAdd) =>
        book.AddFont(fontToAdd);

    public void Close(bool saveChanges, string filename) =>
        book.Close(saveChanges, filename);

    public void Close(bool saveChanges) =>
        book.Close(saveChanges);

    public void Close() =>
        book.Close();

    public void SaveAs(string filename) =>
        book.SaveAs(filename);

    public void SaveAsXml(string filename, ExcelXmlSaveType type) =>
        book.SaveAsXml(filename, type);

    public void SaveAs(string fileName, string separator) =>
        book.SaveAs(fileName, separator);

    public void SaveAs(string fileName, string separator, Encoding encoding) =>
        book.SaveAs(fileName, separator, encoding);

    public void SaveAsHtml(string filename) =>
        book.SaveAsHtml(filename);

    public void SaveAsHtml(string filename, HtmlSaveOptions saveOptions) =>
        book.SaveAsHtml(filename, saveOptions);

    public void SaveAsHtml(Stream stream) =>
        book.SaveAsHtml(stream);

    public void SaveAsHtml(Stream stream, HtmlSaveOptions saveOptions) =>
        book.SaveAsHtml(stream, saveOptions);

    public IFont CreateFont(Font nativeFont) =>
        book.CreateFont(nativeFont);

    public void Replace(string oldValue, DataTable newValues, bool isFieldNamesShown) =>
        book.Replace(oldValue, newValues, isFieldNamesShown);

    public void Replace(string oldValue, DataColumn newValues, bool isFieldNamesShown) =>
        book.Replace(oldValue, newValues, isFieldNamesShown);

    public IHFEngine CreateHFEngine() =>
        book.CreateHFEngine();

    public ITemplateMarkersProcessor CreateTemplateMarkersProcessor() =>
        book.CreateTemplateMarkersProcessor();

    public void MarkAsFinal() =>
        book.MarkAsFinal();

    public void SaveAs(Stream stream) =>
        book.SaveAs(stream);

    public void SaveAs(Stream stream, ExcelSaveType saveType) =>
        book.SaveAs(stream, saveType);

    public void SaveAsXml(Stream stream, ExcelXmlSaveType saveType) =>
        book.SaveAsXml(stream, saveType);

    public void SaveAs(Stream stream, string separator) =>
        book.SaveAs(stream, separator);

    public void SaveAs(Stream stream, string separator, Encoding encoding) =>
        book.SaveAs(stream, separator, encoding);

    public void SaveAsJson(string filename) =>
        book.SaveAsJson(filename);

    public void SaveAsJson(string filename, bool isSchema) =>
        book.SaveAsJson(filename, isSchema);

    public void SaveAsJson(string filename, Sheet worksheet) =>
        book.SaveAsJson(filename, worksheet);

    public void SaveAsJson(string filename, Sheet worksheet, bool isSchema) =>
        book.SaveAsJson(filename, worksheet, isSchema);

    public void SaveAsJson(string filename, Range range) =>
        book.SaveAsJson(filename, range);

    public void SaveAsJson(string filename, Range range, bool isSchema) =>
        book.SaveAsJson(filename, range, isSchema);

    public void SaveAsJson(Stream stream) =>
        book.SaveAsJson(stream);

    public void SaveAsJson(Stream stream, bool isSchema) =>
        book.SaveAsJson(stream, isSchema);

    public void SaveAsJson(Stream stream, Sheet worksheet) =>
        book.SaveAsJson(stream, worksheet);

    public void SaveAsJson(Stream stream, Sheet worksheet, bool isSchema) =>
        book.SaveAsJson(stream, worksheet, isSchema);

    public void SaveAsJson(Stream stream, Range range) =>
        book.SaveAsJson(stream, range);

    public void SaveAsJson(Stream stream, Range range, bool isSchema) =>
        book.SaveAsJson(stream, range, isSchema);

    public void SetPaletteColor(int index, Color color) =>
        book.SetPaletteColor(index, color);

    public void ResetPalette() =>
        book.ResetPalette();

    public Color GetPaletteColor(ExcelKnownColors color) =>
        book.GetPaletteColor(color);

    public ExcelKnownColors GetNearestColor(Color color) =>
        book.GetNearestColor(color);

    public ExcelKnownColors GetNearestColor(int r, int g, int b) =>
        book.GetNearestColor(r, g, b);

    public ExcelKnownColors SetColorOrGetNearest(Color color) =>
        book.SetColorOrGetNearest(color);

    public ExcelKnownColors SetColorOrGetNearest(int r, int g, int b) =>
        book.SetColorOrGetNearest(r, g, b);

    public IFont CreateFont() =>
        book.CreateFont();

    public IFont CreateFont(IFont baseFont) =>
        book.CreateFont(baseFont);

    public void Replace(string oldValue, string newValue) =>
        book.Replace(oldValue, newValue);

    public void Replace(string oldValue, string newValue, ExcelFindOptions findOptions) =>
        book.Replace(oldValue, newValue, findOptions);

    public void Replace(string oldValue, double newValue) =>
        book.Replace(oldValue, newValue);

    public void Replace(string oldValue, DateTime newValue) =>
        book.Replace(oldValue, newValue);

    public void Replace(string oldValue, string[] newValues, bool isVertical) =>
        book.Replace(oldValue, newValues, isVertical);

    public void Replace(string oldValue, int[] newValues, bool isVertical) =>
        book.Replace(oldValue, newValues, isVertical);

    public void Replace(string oldValue, double[] newValues, bool isVertical) =>
        book.Replace(oldValue, newValues, isVertical);

    public Range FindFirst(string findValue, ExcelFindType flags) =>
        book.FindFirst(findValue, flags);

    public Range FindFirst(string findValue, ExcelFindType flags, ExcelFindOptions findOptions) =>
        book.FindFirst(findValue, flags, findOptions);

    public Range FindStringStartsWith(string findValue, ExcelFindType flags) =>
        book.FindStringStartsWith(findValue, flags);

    public Range FindStringStartsWith(string findValue, ExcelFindType flags, bool ignoreCase) =>
        book.FindStringStartsWith(findValue, flags, ignoreCase);

    public Range FindStringEndsWith(string findValue, ExcelFindType flags) =>
        book.FindStringEndsWith(findValue, flags);

    public Range FindStringEndsWith(string findValue, ExcelFindType flags, bool ignoreCase) =>
        book.FindStringEndsWith(findValue, flags, ignoreCase);

    public Range FindFirst(double findValue, ExcelFindType flags) =>
        book.FindFirst(findValue, flags);

    public Range FindFirst(bool findValue) =>
        book.FindFirst(findValue);

    public Range FindFirst(DateTime findValue) =>
        book.FindFirst(findValue);

    public Range FindFirst(TimeSpan findValue) =>
        book.FindFirst(findValue);

    public Range[] FindAll(string findValue, ExcelFindType flags) =>
        book.FindAll(findValue, flags);

    public Range[] FindAll(string findValue, ExcelFindType flags, ExcelFindOptions findOptions) =>
        book.FindAll(findValue, flags, findOptions);

    public Range[] FindAll(double findValue, ExcelFindType flags) =>
        book.FindAll(findValue, flags);

    public Range[] FindAll(bool findValue) =>
        book.FindAll(findValue);

    public Range[] FindAll(DateTime findValue) =>
        book.FindAll(findValue);

    public Range[] FindAll(TimeSpan findValue) =>
        book.FindAll(findValue);

    public void SetSeparators(char argumentsSeparator, char arrayRowsSeparator) =>
        book.SetSeparators(argumentsSeparator, arrayRowsSeparator);

    public void Protect(bool bIsProtectWindow, bool bIsProtectContent) =>
        book.Protect(bIsProtectWindow, bIsProtectContent);

    public void Protect(bool bIsProtectWindow, bool bIsProtectContent, string password) =>
        book.Protect(bIsProtectWindow, bIsProtectContent, password);

    public void Unprotect() =>
        book.Unprotect();

    public void Unprotect(string password) =>
        book.Unprotect(password);

    public Book Clone() =>
        book.Clone();

    public void SetWriteProtectionPassword(string password) =>
        book.SetWriteProtectionPassword(password);

    public void ImportXml(Stream stream) =>
        book.ImportXml(stream);

    public IVbaProject VbaProject => book.VbaProject;

    public ITableStyles TableStyles => book.TableStyles;

    public Sheet ActiveSheet => book.ActiveSheet;

    public int ActiveSheetIndex
    {
        get => book.ActiveSheetIndex;
        set => book.ActiveSheetIndex = value;
    }

    public IAddInFunctions AddInFunctions => book.AddInFunctions;

    public string Author
    {
        get => book.Author;
        set => book.Author = value;
    }

    public bool IsHScrollBarVisible
    {
        get => book.IsHScrollBarVisible;
        set => book.IsHScrollBarVisible = value;
    }

    public bool IsVScrollBarVisible
    {
        get => book.IsVScrollBarVisible;
        set => book.IsVScrollBarVisible = value;
    }

    public IBuiltInDocumentProperties BuiltInDocumentProperties => book.BuiltInDocumentProperties;

    public string CodeName
    {
        get => book.CodeName;
        set => book.CodeName = value;
    }

    public ICustomDocumentProperties CustomDocumentProperties => book.CustomDocumentProperties;

    public IMetaProperties ContentTypeProperties => book.ContentTypeProperties;

    public ICustomXmlPartCollection CustomXmlparts => book.CustomXmlparts;

    public bool Date1904
    {
        get => book.Date1904;
        set => book.Date1904 = value;
    }

    public bool PrecisionAsDisplayed
    {
        get => book.PrecisionAsDisplayed;
        set => book.PrecisionAsDisplayed = value;
    }

    public bool IsCellProtection => book.IsCellProtection;

    public bool IsWindowProtection => book.IsWindowProtection;

    public INames Names => book.Names;

    public bool ReadOnly => book.ReadOnly;

    public bool Saved
    {
        get => book.Saved;
        set => book.Saved = value;
    }

    public IStyles Styles => book.Styles;

    public IWorksheets Worksheets => book.Worksheets;

    public bool HasMacros => book.HasMacros;

    [Obsolete("IWorkbook.Palettte property is obsolete so please use the IWorkbook.Palette property instead. IWorkbook.Palettte will be removed in July 2006. Sorry for the inconvenience")]
    public Color[] Palettte => book.Palettte;

    public Color[] Palette => book.Palette;

    public int DisplayedTab
    {
        get => book.DisplayedTab;
        set => book.DisplayedTab = value;
    }

    public ICharts Charts => book.Charts;

    public bool ThrowOnUnknownNames
    {
        get => book.ThrowOnUnknownNames;
        set => book.ThrowOnUnknownNames = value;
    }

    public bool DisableMacrosStart
    {
        get => book.DisableMacrosStart;
        set => book.DisableMacrosStart = value;
    }

    public double StandardFontSize
    {
        get => book.StandardFontSize;
        set => book.StandardFontSize = value;
    }

    public string StandardFont
    {
        get => book.StandardFont;
        set => book.StandardFont = value;
    }

    public bool Allow3DRangesInDataValidation
    {
        get => book.Allow3DRangesInDataValidation;
        set => book.Allow3DRangesInDataValidation = value;
    }

    public ICalculationOptions CalculationOptions => book.CalculationOptions;

    public string RowSeparator => book.RowSeparator;

    public string ArgumentsSeparator => book.ArgumentsSeparator;

    public IWorksheetGroup WorksheetGroup => book.WorksheetGroup;

    public bool IsRightToLeft
    {
        get => book.IsRightToLeft;
        set => book.IsRightToLeft = value;
    }

    public bool DisplayWorkbookTabs
    {
        get => book.DisplayWorkbookTabs;
        set => book.DisplayWorkbookTabs = value;
    }

    public ITabSheets TabSheets => book.TabSheets;

    public bool DetectDateTimeInValue
    {
        get => book.DetectDateTimeInValue;
        set => book.DetectDateTimeInValue = value;
    }

    public bool UseFastStringSearching
    {
        get => book.UseFastStringSearching;
        set => book.UseFastStringSearching = value;
    }

    public bool ReadOnlyRecommended
    {
        get => book.ReadOnlyRecommended;
        set => book.ReadOnlyRecommended = value;
    }

    public string PasswordToOpen
    {
        get => book.PasswordToOpen;
        set => book.PasswordToOpen = value;
    }

    public int MaxRowCount => book.MaxRowCount;

    public int MaxColumnCount => book.MaxColumnCount;

    public ExcelVersion Version
    {
        get => book.Version;
        set => book.Version = value;
    }

    public IPivotCaches PivotCaches => book.PivotCaches;

    public IConnections Connections => book.Connections;

    public XmlMapCollection XmlMaps => book.XmlMaps;

    public event EventHandler? OnFileSaved
    {
        add => book.OnFileSaved += value;
        remove => book.OnFileSaved -= value;
    }

    public event ReadOnlyFileEventHandler? OnReadOnlyFile
    {
        add => book.OnReadOnlyFile += value;
        remove => book.OnReadOnlyFile -= value;
    }
}