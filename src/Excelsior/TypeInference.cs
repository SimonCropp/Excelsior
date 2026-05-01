namespace Excelsior;

static class TypeInference
{
    static readonly IReadOnlyList<string> boolAllowedValues = ["TRUE", "FALSE"];

    /// <summary>
    /// Returns the auto-derived dropdown list for a column type, or <c>null</c>. Enum columns yield
    /// the rendered enum members; <c>bool</c> / <c>bool?</c> columns yield <c>TRUE</c> / <c>FALSE</c>.
    /// Always-on: the values match what the renderer writes to cells, so the dropdown safely
    /// constrains edits without rejecting legitimate written values.
    /// </summary>
    public static IReadOnlyList<string>? DeriveAllowedValues(Type type)
    {
        var underlying = Nullable.GetUnderlyingType(type) ?? type;
        if (underlying == typeof(bool))
        {
            return boolAllowedValues;
        }

        if (!underlying.IsEnum)
        {
            return null;
        }

        var (_, render) = ValueRenderer.GetRender(underlying);
        var values = Enum.GetValues(underlying);
        var list = new List<string>(values.Length);
        foreach (var value in values)
        {
            list.Add(render?.Invoke(value!) ?? value!.ToString()!);
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
