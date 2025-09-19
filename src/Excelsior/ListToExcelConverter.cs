namespace Excelsior;

/// <summary>
/// Generic converter to export lists to Excel with configurable column styling
/// </summary>
public class ListToExcelConverter<T>(
    List<T> data,
    Action<IXLStyle>? headerStyle = null)
    where T : class
{
    static ListToExcelConverter() =>
        properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(_ => _.CanRead)
            .ToList();

    Dictionary<string, ColumnSettings> columnConfigurations = new();
    ExcelConfiguration excelConfiguration = new();
    static IReadOnlyList<PropertyInfo> properties;

    /// <summary>
    /// Configure a specific column's appearance and behavior
    /// </summary>
    public ListToExcelConverter<T> ConfigureColumn(string property, Action<ColumnSettings> configuration)
    {
        var config = new ColumnSettings();
        configuration(config);
        columnConfigurations[property] = config;
        return this;
    }

    /// <summary>
    /// Configure a column using property expression (type-safe)
    /// </summary>
    /// <returns>The converter instance for fluent chaining</returns>
    public ListToExcelConverter<T> ConfigureColumn<TProperty>(
        Expression<Func<T, TProperty>> property,
        Action<ColumnSettings> configuration)
    {
        var propertyName = GetPropertyName(property);
        return ConfigureColumn(propertyName, configuration);
    }

    /// <returns>The converter instance for fluent chaining</returns>
    public ListToExcelConverter<T> ConfigureExcel(Action<ExcelConfiguration> configuration)
    {
        configuration(excelConfiguration);
        return this;
    }

    public void ExportToStream(Stream stream)
    {
        using var workbook = CreateWorkbook();
        workbook.SaveAs(stream);
    }

    public XLWorkbook CreateWorkbook()
    {
        var workbook = new XLWorkbook();
        AddSheet(workbook);
        return workbook;
    }

    internal void AddSheet(XLWorkbook workbook)
    {
        var worksheet = workbook.Worksheets.Add(excelConfiguration.WorksheetName);

        var properties = GetProperties();

        CreateHeaders(worksheet, properties);

        PopulateData(worksheet, properties);

        ApplyGlobalStyling(worksheet, properties);
        worksheet.RangeUsed()!.SetAutoFilter();
        AutoSizeColumns(worksheet, properties);
    }

    List<PropertyInfo> GetProperties() =>
        // Order by display order if specified
        properties
            .OrderBy(_ =>
            {
                var config = columnConfigurations.GetValueOrDefault(_.Name);
                return config?.Order ?? GetDisplayOrder(_) ?? int.MaxValue;
            })
            .ToList();

    void CreateHeaders(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var cell = worksheet.Cell(1, i + 1);

            // Set header text
            var headerText = GetHeaderText(property);
            cell.Value = headerText;

            ApplyHeaderStyling(cell, property);
        }

        worksheet.SheetView.FreezeRows(1);
    }

    void PopulateData(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        //Skip header
        var startRow = 2;

        for (var rowIndex = 0; rowIndex < data.Count; rowIndex++)
        {
            var item = data[rowIndex];
            var xlRow = startRow + rowIndex;

            for (var colIndex = 0; colIndex < properties.Count; colIndex++)
            {
                var property = properties[colIndex];
                var cell = worksheet.Cell(xlRow, colIndex + 1);

                var value = property.GetValue(item);
                SetCellValue(cell, value, property);
                ApplyDataCellStyling(cell, property, rowIndex);
            }
        }
    }

    void SetCellValue(IXLCell cell, object? value, PropertyInfo property)
    {
        if (value == null)
        {
            var config = columnConfigurations.GetValueOrDefault(property.Name);
            cell.Value = config?.NullDisplayText ?? "";
        }
        else
        {
            var config = columnConfigurations.GetValueOrDefault(property.Name);

            // Apply custom formatter if provided
            if (config?.CustomFormatter != null)
            {
                cell.Value = config.CustomFormatter(value);
                return;
            }

            // Handle specific types
            switch (value)
            {
                case DateTime dateTime:
                    cell.Value = dateTime;
                    if (!string.IsNullOrEmpty(config?.DateTimeFormat))
                        cell.Style.DateFormat.Format = config.DateTimeFormat;
                    break;

                case decimal decimalValue:
                    cell.Value = decimalValue;
                    if (!string.IsNullOrEmpty(config?.NumberFormat))
                        cell.Style.NumberFormat.Format = config.NumberFormat;
                    break;
                case double doubleValue:
                    cell.Value = doubleValue;
                    if (!string.IsNullOrEmpty(config?.NumberFormat))
                        cell.Style.NumberFormat.Format = config.NumberFormat;
                    break;
                case float floatValue:
                    cell.Value = floatValue;
                    if (!string.IsNullOrEmpty(config?.NumberFormat))
                        cell.Style.NumberFormat.Format = config.NumberFormat;
                    break;

                case bool boolean:
                    cell.Value = config?.BooleanDisplayFormat?.Invoke(boolean) ?? boolean.ToString();
                    break;

                case Enum enumValue:
                    cell.Value = config?.EnumDisplayFormat?.Invoke(enumValue) ?? GetEnumDisplayText(enumValue);
                    break;

                default:
                    cell.Value = value.ToString();
                    break;
            }
        }
    }

    void ApplyHeaderStyling(IXLCell cell, PropertyInfo property)
    {
        var config = columnConfigurations.GetValueOrDefault(property.Name);

        // Apply global header styling
        headerStyle?.Invoke(cell.Style);

        // Apply column-specific header styling
        config?.HeaderStyle?.Invoke(cell.Style);
    }

    void ApplyDataCellStyling(IXLCell cell, PropertyInfo property, int rowIndex)
    {
        var config = columnConfigurations.GetValueOrDefault(property.Name);

        // Apply alternating row colors
        if (excelConfiguration.UseAlternatingRowColors && rowIndex % 2 == 1)
        {
            cell.Style.Fill.BackgroundColor = excelConfiguration.AlternateRowColor;
        }

        // Apply column-specific data styling
        config?.DataCellStyle?.Invoke(cell.Style);

        // Apply conditional styling
        if (config?.ConditionalStyling != null)
        {
            var value = property.GetValue(data[rowIndex]);
            config.ConditionalStyling(cell.Style, value);
        }
    }

    void ApplyGlobalStyling(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        if (excelConfiguration.GlobalStyle == null)
        {
            return;
        }

        var range = worksheet.Range(1, 1, data.Count + 1, properties.Count);
        excelConfiguration.GlobalStyle(range.Style);
    }

    void AutoSizeColumns(IXLWorksheet worksheet, List<PropertyInfo> properties)
    {
        worksheet.Columns().AdjustToContents();

        // Apply specific column widths
        for (var i = 0; i < properties.Count; i++)
        {
            var property = properties[i];
            var config = columnConfigurations.GetValueOrDefault(property.Name);

            if (config?.ColumnWidth.HasValue == true)
            {
                worksheet.Column(i + 1).Width = config.ColumnWidth.Value;
            }
        }
    }

    string GetHeaderText(PropertyInfo property)
    {
        var config = columnConfigurations.GetValueOrDefault(property.Name);
        if (!string.IsNullOrEmpty(config?.HeaderText))
            return config.HeaderText;

        // Check for DisplayAttribute
        var displayAttr = property.GetCustomAttribute<DisplayAttribute>();
        if (!string.IsNullOrEmpty(displayAttr?.Name))
        {
            return displayAttr.Name;
        }

        // Check for DisplayNameAttribute
        var displayNameAttr = property.GetCustomAttribute<DisplayNameAttribute>();
        if (!string.IsNullOrEmpty(displayNameAttr?.DisplayName))
        {
            return displayNameAttr.DisplayName;
        }

        // Use property name with spaces
        return AddSpacesToCamelCase(property.Name);
    }

    static int? GetDisplayOrder(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Order;
    }

    static string GetEnumDisplayText(Enum enumValue)
    {
        var field = enumValue.GetType().GetField(enumValue.ToString());
        var attribute = field?.GetCustomAttribute<DisplayAttribute>();
        return attribute?.Name ?? enumValue.ToString();
    }

    static string GetPropertyName<TProperty>(Expression<Func<T, TProperty>> propertyExpression)
    {
        if (propertyExpression.Body is MemberExpression member)
        {
            return member.Member.Name;
        }

        throw new ArgumentException("Expression must be a property access", nameof(propertyExpression));
    }

    static string AddSpacesToCamelCase(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var result = new StringBuilder();
        for (var i = 0; i < text.Length; i++)
        {
            if (i > 0 && char.IsUpper(text[i]))
            {
                result.Append(' ');
            }

            result.Append(text[i]);
        }

        return result.ToString();
    }
}