static class Extensions
{
    public static T? Attribute<T>(this MemberInfo? element)
        where T : Attribute =>
        element?.GetCustomAttribute<T>();

    public static T? Attribute<T>(this ParameterInfo? element)
        where T : Attribute =>
        element?.GetCustomAttribute<T>();

    public static string PropertyName<T, TProperty>(this Expression<Func<T, TProperty>> expression)
    {
        if (expression.Body is MemberExpression member)
        {
            return member.Member.Name;
        }

        throw new ArgumentException("Expression must be a property access", nameof(expression));
    }

    public static string DisplayName(this Enum enumValue)
    {
        var field = enumValue.GetType().GetField(enumValue.ToString());
        var attribute = field?.Attribute<DisplayAttribute>();
        return attribute?.Name ?? enumValue.ToString();
    }

    public static bool IsNumericType(this Type type)
    {
        if (type.IsEnum)
        {
            return false;
        }

        return Type.GetTypeCode(type)
            switch
            {
                TypeCode.Byte or
                    TypeCode.SByte or
                    TypeCode.UInt16 or
                    TypeCode.UInt32 or
                    TypeCode.UInt64 or
                    TypeCode.Int16 or
                    TypeCode.Int32 or
                    TypeCode.Int64 or
                    TypeCode.Decimal or
                    TypeCode.Double or
                    TypeCode.Single => true,
                _ => false
            };
    }
}