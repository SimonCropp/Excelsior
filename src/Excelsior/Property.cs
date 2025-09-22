namespace Excelsior;

class Property
{
    public Property(PropertyInfo info)
    {
        Info = info;
        Name = info.Name;

        var attribute = info.GetCustomAttribute<DisplayAttribute>();
        Order = attribute?.Order;
        DisplayName = GetDisplayName(info);
    }

    public string DisplayName { get; }

    public string Name { get; }

    public int? Order { get; }

    public PropertyInfo Info { get; }

    static string GetDisplayName(PropertyInfo info)
    {
        var display = info.GetCustomAttribute<DisplayAttribute>();
        if (display?.Name != null)
        {
            return display.Name;
        }

        var displayName = info.GetCustomAttribute<DisplayNameAttribute>();
        if (displayName != null)
        {
            return displayName.DisplayName;
        }

        return CamelCase.Split(info.Name);
    }
}