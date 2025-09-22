static class Extensions
{
    public static string PropertyName<T, TProperty>(this Expression<Func<T, TProperty>> propertyExpression)
    {
        if (propertyExpression.Body is MemberExpression member)
        {
            return member.Member.Name;
        }

        throw new ArgumentException("Expression must be a property access", nameof(propertyExpression));
    }

    public static bool IsNumericType(this Type type) =>
        Type.GetTypeCode(type)
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