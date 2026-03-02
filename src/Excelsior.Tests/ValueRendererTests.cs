[TestFixture]
public class ValueRendererTests
{
    public interface IAnimal
    {
        string Name { get; }
    }

    public record Dog(string Name) : IAnimal;

    public enum TestColor
    {
        Red,
        Blue
    }

    [ModuleInitializer]
    public static void Init()
    {
        ValueRenderer.For<IAnimal>(_ => $"Animal: {_.Name}");
        ValueRenderer.NullDisplayFor<IAnimal>("No Animal");
        ValueRenderer.For<TestColor>(_ => $"Color: {_}");
        ValueRenderer.NullDisplayFor<TestColor>("No Color");
    }

    [Test]
    public void GetRender_WithSubtype_ReturnsRender()
    {
        var (isEnumerable, render) = ValueRenderer.GetRender(typeof(Dog));

        Assert.That(render, Is.Not.Null);
        Assert.That(isEnumerable, Is.False);
        Assert.That(render!(new Dog("Rex")), Is.EqualTo("Animal: Rex"));
    }

    [Test]
    public void GetRender_WithExactType_ReturnsRender()
    {
        var (isEnumerable, render) = ValueRenderer.GetRender(typeof(IAnimal));

        Assert.That(render, Is.Not.Null);
        Assert.That(isEnumerable, Is.False);
    }

    [Test]
    public void GetRender_WithUnrelatedType_ReturnsNull()
    {
        var (_, render) = ValueRenderer.GetRender(typeof(string));

        Assert.That(render, Is.Null);
    }

    [Test]
    public void GetRender_WithNullableValueType_ReturnsRender()
    {
        var (isEnumerable, render) = ValueRenderer.GetRender(typeof(TestColor?));

        Assert.That(render, Is.Not.Null);
        Assert.That(isEnumerable, Is.False);
        Assert.That(render!(TestColor.Red), Is.EqualTo("Color: Red"));
    }

    [Test]
    public void GetNullDisplay_WithSubtype_ReturnsDisplay()
    {
        var result = ValueRenderer.GetNullDisplay(typeof(Dog));

        Assert.That(result, Is.EqualTo("No Animal"));
    }

    [Test]
    public void GetNullDisplay_WithExactType_ReturnsDisplay()
    {
        var result = ValueRenderer.GetNullDisplay(typeof(IAnimal));

        Assert.That(result, Is.EqualTo("No Animal"));
    }

    [Test]
    public void GetNullDisplay_WithUnrelatedType_ReturnsNull()
    {
        var result = ValueRenderer.GetNullDisplay(typeof(string));

        Assert.That(result, Is.Null);
    }

    [Test]
    public void GetNullDisplay_WithNullableValueType_ReturnsDisplay()
    {
        var result = ValueRenderer.GetNullDisplay(typeof(TestColor?));

        Assert.That(result, Is.EqualTo("No Color"));
    }

    [Test]
    public void GetRender_EnumerableOfSubtype_ReturnsRender()
    {
        var (isEnumerable, render) = ValueRenderer.GetRender(typeof(IReadOnlyList<Dog>));

        Assert.That(isEnumerable, Is.True);
        Assert.That(render, Is.Not.Null);
    }

    [Test]
    public void GetRender_EnumerableOfExactType_ReturnsRender()
    {
        var (isEnumerable, render) = ValueRenderer.GetRender(typeof(IReadOnlyList<IAnimal>));

        Assert.That(isEnumerable, Is.True);
        Assert.That(render, Is.Not.Null);
    }
}
