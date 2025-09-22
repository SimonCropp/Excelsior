namespace Excelsior;

class Property
{
    public Property(PropertyInfo info)
    {
        Info = info;

        var attribute = info.GetCustomAttribute<DisplayAttribute>();
        Order =  attribute?.Order;
    }

    public int? Order { get; set; }

    public PropertyInfo Info { get; }
}