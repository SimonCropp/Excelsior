[TestFixture]
public class EnumExtensionsTests
{
    [Test]
    public void Humanize_PascalCase_AddsSpacesAndLowercases()
    {
        var result = TestEnum.AntiqueWhite.Humanize();

        Assert.That(result, Is.EqualTo("Antique white"));
    }

    [Test]
    public void Humanize_MultipleCamelCaseWords_AddsSpacesBetweenEach()
    {
        var result = TestEnum.DeepSkyBlue.Humanize();

        Assert.That(result, Is.EqualTo("Deep sky blue"));
    }

    [Test]
    public void Humanize_SingleWord_ReturnsAsIs()
    {
        var result = TestEnum.Red.Humanize();

        Assert.That(result, Is.EqualTo("Red"));
    }

    [Test]
    public void Humanize_WithDisplayAttribute_ReturnsDisplayName()
    {
        var result = TestEnum.CustomColor.Humanize();

        Assert.That(result, Is.EqualTo("My Custom Color"));
    }

    [Test]
    public void Humanize_WithEmptyDisplayAttribute_FallsBackToHumanization()
    {
        var result = TestEnum.FallbackName.Humanize();

        Assert.That(result, Is.EqualTo("Fallback name"));
    }

    [Test]
    public void Humanize_CalledTwice_ReturnsSameInstance()
    {
        var first = TestEnum.AntiqueWhite.Humanize();
        var second = TestEnum.AntiqueWhite.Humanize();

        Assert.That(ReferenceEquals(first, second), Is.True, "Should return cached instance");
    }

    [Test]
    public void Humanize_DifferentEnumTypes_DoNotCollide()
    {
        var colorResult = TestEnum.Red.Humanize();
        var shapeResult = AnotherEnum.Red.Humanize();

        Assert.That(colorResult, Is.EqualTo("Red"));
        Assert.That(shapeResult, Is.EqualTo("Red"));
    }

    [Test]
    public void Humanize_AllUppercase_ReturnsUnchanged()
    {
        var result = TestEnum.RGB.Humanize();

        Assert.That(result, Is.EqualTo("RGB"));
    }

    [Test]
    public void Humanize_AllUppercaseAcronym_ReturnsUnchanged()
    {
        var result = TestEnum.HTML.Humanize();

        Assert.That(result, Is.EqualTo("HTML"));
    }

    [Test]
    public void Humanize_MixedWithAcronym_AddsSpaces()
    {
        var result = TestEnum.HTMLColor.Humanize();

        Assert.That(result, Is.EqualTo("H t m l color"));
    }

    [Test]
    [TestCase(TestEnum.Red)]
    [TestCase(TestEnum.AntiqueWhite)]
    [TestCase(TestEnum.DeepSkyBlue)]
    public void Humanize_MultipleValues_EachCachedIndependently(TestEnum value)
    {
        var result = value.Humanize();

        Assert.That(result, Is.Not.Null);
        Assert.That(result, Is.Not.Empty);
    }

    [Test]
    public void Humanize_ThreadSafety_NoExceptions()
    {
        var exceptions = new List<Exception>();
        var threads = new List<Thread>();

        for (var i = 0; i < 10; i++)
        {
            var thread = new Thread(() =>
            {
                try
                {
                    for (var j = 0; j < 100; j++)
                    {
                        TestEnum.AntiqueWhite.Humanize();
                        TestEnum.DeepSkyBlue.Humanize();
                        TestEnum.CustomColor.Humanize();
                    }
                }
                catch (Exception ex)
                {
                    lock (exceptions)
                    {
                        exceptions.Add(ex);
                    }
                }
            });
            threads.Add(thread);
            thread.Start();
        }

        foreach (var thread in threads)
        {
            thread.Join();
        }

        Assert.That(exceptions, Is.Empty, $"Exceptions occurred: {string.Join(", ", exceptions.Select(e => e.Message))}");
    }

    [Test]
    public void Humanize_StartsWithLowercase_PreservesFirstLetter()
    {
        var result = EdgeCaseEnum.lowercase.Humanize();

        Assert.That(result, Is.EqualTo("lowercase"));
    }

    [Test]
    public void Humanize_AllUppercaseWord_ReturnsUnchanged()
    {
        var result = EdgeCaseEnum.UPPERCASE.Humanize();

        Assert.That(result, Is.EqualTo("UPPERCASE"));
    }

    public enum TestEnum
    {
        Red,
        AntiqueWhite,
        DeepSkyBlue,
        RGB,
        HTML,
        HTMLColor,

        [Display(Name = "My Custom Color")]
        CustomColor,

        [Display]
        FallbackName
    }

    public enum AnotherEnum
    {
        Red,
        Blue
    }

    public enum EdgeCaseEnum
    {
        lowercase,
        UPPERCASE,
        MixedCASEValue
    }
}