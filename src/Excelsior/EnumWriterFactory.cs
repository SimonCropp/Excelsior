internal delegate void TypedCellWriter<TModel>(TModel item, Cell cell, ColumnConfig<TModel> column);

static class EnumWriterFactory<TModel>
{
    static readonly MethodInfo buildMethod = typeof(EnumWriterFactory<TModel>)
        .GetMethod(nameof(BuildEnum), BindingFlags.NonPublic | BindingFlags.Static)!;
    static readonly MethodInfo buildNullableMethod = typeof(EnumWriterFactory<TModel>)
        .GetMethod(nameof(BuildNullableEnum), BindingFlags.NonPublic | BindingFlags.Static)!;

    public static TypedCellWriter<TModel>? TryBuild(Type propType, IEnumerable<MemberInfo> path)
    {
        var underlying = Nullable.GetUnderlyingType(propType);
        var effective = underlying ?? propType;
        if (!effective.IsEnum)
        {
            return null;
        }

        var method = underlying != null
            ? buildNullableMethod.MakeGenericMethod(effective)
            : buildMethod.MakeGenericMethod(effective);
        return (TypedCellWriter<TModel>)method.Invoke(null, [path])!;
    }

    static TypedCellWriter<TModel> BuildEnum<TEnum>(IEnumerable<MemberInfo> path)
        where TEnum : struct, Enum
    {
        var getter = Properties<TModel>.CreateTypedGet<TEnum>(path);
        return (item, cell, column) =>
        {
            var value = getter(item);
            var text = EnumRender<TEnum>.Render(value);
            CellWrite.StringOrHtml(cell, text, column.IsHtml);
        };
    }

    static TypedCellWriter<TModel> BuildNullableEnum<TEnum>(IEnumerable<MemberInfo> path)
        where TEnum : struct, Enum
    {
        var getter = Properties<TModel>.CreateTypedGet<TEnum?>(path);
        return (item, cell, column) =>
        {
            var value = getter(item);
            if (value == null)
            {
                if (column.NullDisplay != null)
                {
                    CellWrite.String(cell, column.NullDisplay);
                }

                return;
            }

            var text = EnumRender<TEnum>.Render(value.Value);
            CellWrite.StringOrHtml(cell, text, column.IsHtml);
        };
    }
}
