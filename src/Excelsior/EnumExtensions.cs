public static class EnumExtensions
{
    private static readonly ConcurrentDictionary<Enum, string> Cache = new();

    public static string Humanize(this Enum value) =>
        Cache.GetOrAdd(value, static v =>
        {
            var type = v.GetType();
            var memberInfo = type.GetField(v.ToString());

            if (memberInfo is null)
                return v.ToString();

            // Check for DisplayAttribute
            var displayAttribute = memberInfo.GetCustomAttribute<DisplayAttribute>();
            if (displayAttribute?.Name is not null)
                return displayAttribute.Name;

            // Humanize the enum name
            return HumanizeName(v.ToString());
        });

    private static string HumanizeName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return name;

        // If all characters are uppercase, return as-is
        if (name.All(char.IsUpper))
            return name;

        var builder = new StringBuilder(name.Length + 5);
        builder.Append(name[0]);

        for (var i = 1; i < name.Length; i++)
        {
            var current = name[i];

            if (char.IsUpper(current))
            {
                // Add space before uppercase letter
                builder.Append(' ');
                builder.Append(char.ToLowerInvariant(current));
            }
            else
            {
                builder.Append(current);
            }
        }

        return builder.ToString();
    }
}