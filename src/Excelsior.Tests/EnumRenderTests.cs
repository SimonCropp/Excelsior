[TestFixture]
public class EnumRenderTests
{
    [Test]
    public void Render_Default_ReturnsHumanized()
    {
        var result = EnumRender.Render(DefaultColor.DeepSkyBlue);

        Assert.That(result, Is.EqualTo("Deep sky blue"));
    }

    [Test]
    public void Render_Default_HonoursDisplayAttribute()
    {
        var result = EnumRender.Render(DisplayAttrEnum.WithDescription);

        Assert.That(result, Is.EqualTo("This is the description"));
    }

    [Test]
    public void Render_TypedSetOverride_IsHonoured()
    {
        EnumRender<TypedOverrideEnum>.Set(static value => value switch
        {
            TypedOverrideEnum.FullTime => "Full time",
            TypedOverrideEnum.PartTime => "Part time",
            _ => value.ToString()
        });

        var result = EnumRender.Render(TypedOverrideEnum.PartTime);

        Assert.That(result, Is.EqualTo("Part time"));
    }

    [Test]
    public void Render_BoxedNullableEnum_DispatchesToUnderlying()
    {
        // Boxing a Nullable<T> with a value boxes the T itself, so value.GetType()
        // returns the enum type — not Nullable<TEnum>. Guards against MakeGenericType
        // ever being asked for a Nullable<> by the dispatcher cache.
        var value = (NullableDispatchEnum?)NullableDispatchEnum.AntiqueWhite;
        var boxed = (Enum)(object)value!;

        var result = EnumRender.Render(boxed);

        Assert.That(result, Is.EqualTo("Antique white"));
    }

    [Test]
    public void Render_FlagsEnum_FallsBackToHumanize()
    {
        // [Flags] is skipped by the source generator (a value-switch can't represent
        // arbitrary bitwise combinations), so the boxed dispatcher must fall through
        // to the Humanize path for single-value cases.
        var result = EnumRender.Render(FlagsEnum.Bravo);

        Assert.That(result, Is.EqualTo("Bravo"));
    }

    [Test]
    public void Render_RepeatedCalls_ReturnSameCachedDispatcher()
    {
        // Smoke test for the ConcurrentDictionary cache — repeated calls for the same
        // enum type should not allocate a new MakeGenericType / CreateDelegate pair.
        var first = EnumRender.Render(CacheStabilityEnum.One);
        var second = EnumRender.Render(CacheStabilityEnum.One);

        Assert.That(first, Is.EqualTo("One"));
        Assert.That(second, Is.EqualTo("One"));
    }

    enum DefaultColor
    {
        DeepSkyBlue
    }

    enum DisplayAttrEnum
    {
        [Display(Description = "This is the description")]
        WithDescription
    }

    enum TypedOverrideEnum
    {
        FullTime,
        PartTime
    }

    enum NullableDispatchEnum
    {
        AntiqueWhite
    }

    [Flags]
    enum FlagsEnum
    {
        Alpha = 1,
        Bravo = 2,
        Charlie = 4
    }

    enum CacheStabilityEnum
    {
        One
    }
}
