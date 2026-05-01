namespace Excelsior;

class SheetReader<TModel> :
    ISheetReader<TModel>,
    IReaderSheet
{
    Dictionary<string, ColumnReadState> columns = new(StringComparer.Ordinal);
    List<TModel> rows = [];

    public string? Name { get; }
    public IReadOnlyList<TModel> Rows => rows;

    public SheetReader(string? name)
    {
        Name = name;

        foreach (var info in GetReadableProperties(typeof(TModel)))
        {
            columns[info.Name] = new()
            {
                Name = info.Name,
                Heading = ResolveHeading(info),
                Type = info.PropertyType,
                Include = ResolveInclude(info),
                Convert = null
            };
        }
    }

    static IEnumerable<PropertyInfo> GetReadableProperties(Type type)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanRead)
            {
                continue;
            }

            if (property.GetIndexParameters().Length > 0)
            {
                continue;
            }

            if (property.Attribute<IgnoreAttribute>() != null)
            {
                continue;
            }

            yield return property;
        }
    }

    static string ResolveHeading(PropertyInfo info)
    {
        var column = info.Attribute<ColumnAttribute>();
        if (column?.Heading != null)
        {
            return column.Heading;
        }

        var display = info.Attribute<DisplayAttribute>();
        if (display?.Name != null)
        {
            return display.Name;
        }

        var displayName = info.Attribute<DisplayNameAttribute>();
        if (displayName != null)
        {
            return displayName.DisplayName;
        }

        return CamelCase.Split(info.Name);
    }

    static bool ResolveInclude(PropertyInfo info)
    {
        var column = info.Attribute<ColumnAttribute>();
        if (column is { IncludeHasValue: true })
        {
            return column.Include;
        }

        return true;
    }

    public ISheetReader<TModel> Column<TProperty>(
        Expression<Func<TModel, TProperty>> property,
        Action<ColumnReadConfig<TProperty>> configuration)
    {
        var name = property.PropertyName();
        if (!columns.TryGetValue(name, out var column))
        {
            throw new($"Could not find property: {name}");
        }

        var config = new ColumnReadConfig<TProperty>();
        configuration(config);

        if (config.Heading != null)
        {
            column.Heading = config.Heading;
        }

        if (config.Include != null)
        {
            column.Include = config.Include.Value;
        }

        if (config.Convert != null)
        {
            var convert = config.Convert;
            column.Convert = cell => convert(cell);
        }

        return this;
    }

    public void HeadingText<TProperty>(Expression<Func<TModel, TProperty>> property, string value) =>
        Column(property, _ => _.Heading = value);

    public void Include<TProperty>(Expression<Func<TModel, TProperty>> property, bool value) =>
        Column(property, _ => _.Include = value);

    public void Exclude<TProperty>(Expression<Func<TModel, TProperty>> property) =>
        Column(property, _ => _.Include = false);

    public void Convert<TProperty>(Expression<Func<TModel, TProperty>> property, Func<Cell, TProperty> convert) =>
        Column(property, _ => _.Convert = convert);

    IReadOnlyList<ColumnReadInfo> IReaderSheet.Columns()
    {
        var result = new List<ColumnReadInfo>();
        foreach (var c in columns.Values)
        {
            if (c.Include)
            {
                result.Add(new(c.Name, c.Heading, c.Type, c.Convert));
            }
        }

        return result;
    }

    void IReaderSheet.Receive(IReadOnlyDictionary<string, object?> rowValues) =>
        rows.Add(ModelActivator<TModel>.Create(rowValues));

    void IReaderSheet.Reset() =>
        rows.Clear();
}

interface IReaderSheet
{
    string? Name { get; }
    IReadOnlyList<ColumnReadInfo> Columns();
    void Receive(IReadOnlyDictionary<string, object?> rowValues);
    void Reset();
}

sealed record ColumnReadInfo(
    string Name,
    string Heading,
    Type Type,
    Func<Cell, object?>? Convert);
