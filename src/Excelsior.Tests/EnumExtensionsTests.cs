[TestFixture]
public class EnumExtensionsTests
{
    [Test]
    public void Humanize_WithDescriptionOnly_ReturnsDescription()
    {
        var result = TestEnum.WithDescription.Humanize();

        Assert.That(result, Is.EqualTo("This is the description"));
    }

    [Test]
    public void Humanize_WithDescriptionAndName_ReturnsDescription()
    {
        var result = TestEnum.WithBoth.Humanize();

        Assert.That(result, Is.EqualTo("Description wins"));
    }

    [Test]
    public void Humanize_WithNameOnly_ReturnsName()
    {
        var result = TestEnum.WithNameOnly.Humanize();

        Assert.That(result, Is.EqualTo("Custom Name"));
    }

    [Test]
    public void Humanize_WithEmptyDisplayAttribute_FallsBackToHumanization()
    {
        var result = TestEnum.EmptyDisplay.Humanize();

        Assert.That(result, Is.EqualTo("Empty display"));
    }

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
    [TestCase(TestEnum.Red, "Red")]
    [TestCase(TestEnum.AntiqueWhite, "Antique white")]
    [TestCase(TestEnum.DeepSkyBlue, "Deep sky blue")]
    [TestCase(TestEnum.WithDescription, "This is the description")]
    [TestCase(TestEnum.WithBoth, "Description wins")]
    [TestCase(TestEnum.WithNameOnly, "Custom Name")]
    public void Humanize_VariousValues_ReturnsExpectedResult(TestEnum value, string expected)
    {
        var result = value.Humanize();

        Assert.That(result, Is.EqualTo(expected));
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
                        TestEnum.WithDescription.Humanize();
                        TestEnum.WithBoth.Humanize();
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

    [Test]
    public void Humanize_DisplayAttributePriority_DescriptionOverName()
    {
        // Verify that when both Description and Name are set, Description is used
        var result = TestEnum.WithBoth.Humanize();

        Assert.That(result, Is.Not.EqualTo("Name comes second"));
        Assert.That(result, Is.EqualTo("Description wins"));
    }

    public enum TestEnum
    {
        Red,
        AntiqueWhite,
        DeepSkyBlue,
        RGB,
        HTML,
        HTMLColor,

        [Display(Description = "This is the description")]
        WithDescription,

        [Display(Description = "Description wins", Name = "Name comes second")]
        WithBoth,

        [Display(Name = "Custom Name")]
        WithNameOnly,

        [Display]
        EmptyDisplay
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