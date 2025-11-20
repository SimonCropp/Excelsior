public static class EnumExtensions
{
    static ConcurrentDictionary<Enum, string> Cache = new();

    public static string Humanize(this Enum value) =>
        Cache.GetOrAdd(value, static value =>
        {
            var type = value.GetType();
            var memberInfo = type.GetField(value.ToString());

            if (memberInfo is null)
                return value.ToString();

            // Check for DisplayAttribute - Description takes priority over Name
            var displayAttribute = memberInfo.GetCustomAttribute<DisplayAttribute>();
            if (displayAttribute is not null)
            {
                if (displayAttribute.Description is not null)
                {
                    return displayAttribute.Description;
                }

                if (displayAttribute.Name is not null)
                {
                    return displayAttribute.Name;
                }
            }

            // Humanize the enum name
            return HumanizeName(value.ToString());
        });

    static string HumanizeName(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return name;
        }

        // If all characters are uppercase, return as-is
        if (name.All(char.IsUpper))
        {
            return name;
        }

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