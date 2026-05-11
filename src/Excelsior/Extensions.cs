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
            throw new ArgumentException("Expression must be a property or field access", nameof(expression));
        }

        parts.Reverse();
        return string.Join('.', parts);
    }

    public static Type GetMemberType(this MemberInfo member) =>
        member switch
        {
            PropertyInfo property => property.PropertyType,
            FieldInfo field => field.FieldType,
            _ => throw new($"Unsupported member kind: {member.GetType()}")
        };

    public static bool CanReadMember(this MemberInfo member) =>
        member switch
        {
            PropertyInfo property => property.CanRead &&
                                     property.GetIndexParameters().Length == 0,
            FieldInfo => true,
            _ => false
        };

    public static bool CanWriteMember(this MemberInfo member) =>
        member switch
        {
            PropertyInfo property => property.CanWrite && property.GetIndexParameters().Length == 0,
            FieldInfo field => !(field.IsInitOnly || field.IsLiteral),
            _ => false
        };

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
