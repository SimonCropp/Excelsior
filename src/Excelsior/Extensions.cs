static class Extensions
{
    public static T? Attribute<T>(this MemberInfo? element)
        where T : Attribute =>
        element?.GetCustomAttribute<T>();

    public static T? Attribute<T>(this ParameterInfo? element)
        where T : Attribute =>
        element?.GetCustomAttribute<T>();

    public static bool HasAttribute<T>(this ParameterInfo? element)
        where T : Attribute =>
        element?.GetCustomAttribute<T>() != null;

    public static string PropertyName<T, TProperty>(this Expression<Func<T, TProperty>> expression)
    {
        var parts = new List<string>();
        var current = expression.Body;

        while (current is MemberExpression member)
        {
            parts.Add(member.Member.Name);
            current = member.Expression;
        }

        if (parts.Count == 0)
        {
            throw new ArgumentException("Expression must be a property access", nameof(expression));
        }

        parts.Reverse();
        return string.Join('.', parts);
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