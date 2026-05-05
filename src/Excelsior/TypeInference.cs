static class TypeInference
{
    /// <summary>
    /// Returns the auto-derived dropdown list for a column type, or <c>null</c>. For bool columns
    /// the dropdown matches what's displayed in cells: by default <c>TRUE</c>/<c>FALSE</c>, or the
    /// strings supplied via <see cref="ValueRenderer.BoolDisplay"/>. For enums the dropdown is
    /// produced by the same renderer that writes cells, so a custom
    /// <see cref="ValueRenderer.For{T}"/> registration (or the enum humanizer) flows through.
    /// </summary>
    public static IReadOnlyList<string>? DeriveAllowedValues(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type) ?? type;
        if (underlying == typeof(bool))
        {
            var (trueDisplay, falseDisplay) = ValueRenderer.GetBoolDisplayValues();
            return [trueDisplay, falseDisplay];
        }

        if (!underlying.IsEnum)
        {
            return null;
        }

        var (_, enumRender) = ValueRenderer.GetRender(underlying);
        var values = Enum.GetValues(underlying);
        var list = new List<string>(values.Length);
        foreach (var value in values)
        {
            list.Add(enumRender?.Invoke(value!) ?? value!.ToString()!);
        }

        return list;
    }

    /// <summary>
    /// Returns whether a generic <c>TProperty</c> represents a non-nullable value type.
    /// Reference types fall back to <c>false</c> here because nullability annotations are not
    /// reachable from a generic parameter — use a <see cref="System.Reflection.PropertyInfo"/>
    /// for NRT-aware inference.
    /// </summary>
    public static bool IsNonNullableValueType(Type type) =>
        type.IsValueType && Nullable.GetUnderlyingType(type) == null;
}
