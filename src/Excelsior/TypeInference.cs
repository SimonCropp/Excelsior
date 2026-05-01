namespace Excelsior;

static class TypeInference
{
    static readonly IReadOnlyList<string> boolAllowedValues = ["TRUE", "FALSE"];

    /// <summary>
    /// Returns the auto-derived dropdown list for a column type, or <c>null</c>. Always-on: the
    /// values are produced by the same renderer that writes cells, so a custom
    /// <see cref="ValueRenderer.For{T}"/> registration (or the enum humanizer) flows through to the
    /// dropdown — e.g. <c>ValueRenderer.For&lt;bool&gt;(_ =&gt; _ ? "Yes" : "No")</c> yields a
    /// <c>Yes</c> / <c>No</c> dropdown that matches the cell content.
    /// </summary>
    public static IReadOnlyList<string>? DeriveAllowedValues(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type) ?? type;
        if (underlying == typeof(bool))
        {
            var (_, render) = ValueRenderer.GetRender(underlying);
            return render != null
                ? [render(true), render(false)]
                : boolAllowedValues;
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
