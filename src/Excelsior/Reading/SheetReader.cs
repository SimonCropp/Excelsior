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

        if (config.Convert != null)
        {
            var convert = config.Convert;
            column.Convert = cell => convert(cell);
        }

        return this;
    }

    public void HeadingText<TProperty>(Expression<Func<TModel, TProperty>> property, string value) =>
        Column(property, _ => _.Heading = value);

    public void Convert<TProperty>(Expression<Func<TModel, TProperty>> property, Func<Cell, TProperty> convert) =>
        Column(property, _ => _.Convert = convert);

    public List<ColumnReadInfo> Columns()
    {
        var result = new List<ColumnReadInfo>(columns.Count);
        foreach (var column in columns.Values)
        {
            result.Add(new(column.Name, column.Heading, column.Type, column.Convert));
        }

        return result;
    }

    public void Receive(IReadOnlyDictionary<string, object?> rowValues) =>
        rows.Add(ModelActivator<TModel>.Create(rowValues));

    public void Reset() =>
        rows.Clear();
}