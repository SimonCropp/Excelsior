namespace Excelsior;

/// <summary>
/// Per-enum render hook. Each closed generic <c>EnumRender&lt;TEnum&gt;</c> has its own
/// static slot, so the source generator can install a non-boxing switch for each enum
/// it discovers. Resolution order:
/// <list type="number">
///   <item><see cref="ValueRenderer.ForEnums(Func{Enum,string})"/> (global override) wins when set.</item>
///   <item>The per-type <see cref="Set"/> delegate (e.g. from the source generator) is used next.</item>
///   <item>Otherwise falls back to the boxed reflection path in <see cref="EnumExtensions.Humanize"/>.</item>
/// </list>
/// </summary>
public static class EnumRender<TEnum>
    where TEnum : struct, Enum
{
    static Func<TEnum, string>? configured;

    /// <summary>
    /// Install a non-boxing per-type render. Typically called from a
    /// <see cref="System.Runtime.CompilerServices.ModuleInitializerAttribute"/>
    /// emitted by the Excelsior source generator, but available to user code as well.
    /// </summary>
    public static void Set(Func<TEnum, string> render)
    {
        configured = render;
        CellConverter.ResetEnumCache();
    }

    internal static string Render(TEnum value)
    {
        if (ValueRenderer.HasCustomEnumRender)
        {
            return ValueRenderer.RenderEnum(value);
        }

        var c = configured;
        if (c != null)
        {
            return c(value);
        }

        return value.Humanize();
    }
}
