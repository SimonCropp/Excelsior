[TestFixture]
public class PropertiesTests
{
    class SimpleModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";

#pragma warning disable CA1822
        public string WriteOnly
        {
            set => _ = value;
        }
#pragma warning restore CA1822
    }

    class AttributedModel
    {
        [Column(Header = "Custom Header", Order = 1, Width = 150, Format = "N2", NullDisplay = "N/A", IsHtml = true)]
        public decimal Price { get; set; }

        [Display(Name = "Display Name", Order = 2)]
        public string Name { get; set; } = "";

        [DisplayName("DisplayName Attribute")]
        public string Description { get; set; } = "";

        public string NoAttributes { get; set; } = "";
    }

    class OrderTestModel
    {
        [Column(Order = 2)]
        public string Second { get; set; } = "";

        [Display(Order = 1)]
        public string First { get; set; } = "";

        [Column(Order = 3)]
        [Display(Order = 10)]
        public string Third { get; set; } = "";

        public string NoOrder { get; set; } = "";
    }

    class NumericModel
    {
        public int Integer { get; set; }
        public long Long { get; set; }
        public decimal Decimal { get; set; }
        public double Double { get; set; }
        public float Float { get; set; }
        public byte Byte { get; set; }
        public string NotNumeric { get; set; } = "";
    }

    [Test]
    public void Properties_OnlyIncludesReadablePublicInstanceProperties()
    {
        var items = Properties<SimpleModel>.Items;

        Assert.That(items, Has.Count.EqualTo(2));
        Assert.That(items, Has.Some.Matches<Property<SimpleModel>>(p => p.Name == "Id"));
        Assert.That(items, Has.Some.Matches<Property<SimpleModel>>(p => p.Name == "Name"));
        Assert.That(items, Has.None.Matches<Property<SimpleModel>>(p => p.Name == "WriteOnly"));
    }

    [Test]
    public void Properties_CachesResultsAcrossInstances()
    {
        var items1 = Properties<SimpleModel>.Items;
        var items2 = Properties<SimpleModel>.Items;

        Assert.That(items1, Is.SameAs(items2));
    }

    [Test]
    public void Property_Get_ReturnsPropertyValue()
    {
        var model = new SimpleModel
        {
            Id = 42,
            Name = "Test"
        };
        var idProp = Properties<SimpleModel>.Items.First(p => p.Name == "Id");
        var nameProp = Properties<SimpleModel>.Items.First(p => p.Name == "Name");

        Assert.That(idProp.Get(model), Is.EqualTo(42));
        Assert.That(nameProp.Get(model), Is.EqualTo("Test"));
    }

    [Test]
    public void Property_Get_HandlesNullValues()
    {
        var model = new AttributedModel
        {
            Name = null!
        };
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.Get(model), Is.Null);
    }

    [Test]
    public void Property_ColumnAttribute_AllPropertiesRead()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Price");

        Assert.That(prop.DisplayName, Is.EqualTo("Custom Header"));
        Assert.That(prop.Order, Is.EqualTo(1));
        Assert.That(prop.Width, Is.EqualTo(150));
        Assert.That(prop.Format, Is.EqualTo("N2"));
        Assert.That(prop.NullDisplay, Is.EqualTo("N/A"));
        Assert.That(prop.IsHtml, Is.True);
    }

    [Test]
    public void Property_DisplayAttribute_UsedForHeaderAndOrder()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.DisplayName, Is.EqualTo("Display Name"));
        Assert.That(prop.Order, Is.EqualTo(2));
    }

    [Test]
    public void Property_DisplayNameAttribute_UsedForHeader()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Description");

        Assert.That(prop.DisplayName, Is.EqualTo("DisplayName Attribute"));
    }

    [Test]
    public void Property_NoAttributes_UsesCamelCaseSplit()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "NoAttributes");

        Assert.That(prop.DisplayName, Is.EqualTo("No Attributes"));
    }

    [Test]
    public void Property_Order_ColumnAttributeTakesPrecedenceOverDisplay()
    {
        var prop = Properties<OrderTestModel>.Items.First(p => p.Name == "Third");

        Assert.That(prop.Order, Is.EqualTo(3));
    }

    [Test]
    public void Property_Order_NegativeColumnOrderIgnored()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "NoAttributes");

        Assert.That(prop.Order, Is.Null);
    }

    [Test]
    public void Property_Width_NegativeValueReturnsNull()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.Width, Is.Null);
    }

    [Test]
    public void Property_Width_PositiveValueReturned()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Price");

        Assert.That(prop.Width, Is.EqualTo(150));
    }

    [Test]
    public void Property_Format_ReturnsNullWhenNotSet()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.Format, Is.Null);
    }

    [Test]
    public void Property_NullDisplay_ReturnsNullWhenNotSet()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.NullDisplay, Is.Null);
    }

    [Test]
    public void Property_IsHtml_DefaultsToFalse()
    {
        var prop = Properties<AttributedModel>.Items.First(p => p.Name == "Name");

        Assert.That(prop.IsHtml, Is.False);
    }

    [Test]
    public void Property_Type_ReflectsPropertyType()
    {
        var idProp = Properties<SimpleModel>.Items.First(p => p.Name == "Id");
        var nameProp = Properties<SimpleModel>.Items.First(p => p.Name == "Name");

        Assert.That(idProp.Type, Is.EqualTo(typeof(int)));
        Assert.That(nameProp.Type, Is.EqualTo(typeof(string)));
    }

    [Test]
    public void Property_IsNumber_DetectsNumericTypes()
    {
        var props = Properties<NumericModel>.Items;

        Assert.That(props.First(p => p.Name == "Integer").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "Long").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "Decimal").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "Double").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "Float").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "Byte").IsNumber, Is.True);
        Assert.That(props.First(p => p.Name == "NotNumeric").IsNumber, Is.False);
    }

    [Test]
    public void Property_Name_MatchesPropertyName()
    {
        var prop = Properties<SimpleModel>.Items.First(p => p.Name == "Id");

        Assert.That(prop.Name, Is.EqualTo("Id"));
    }

    class ComplexModel
    {
        [Column(Header = "ID", Order = 0, Width = 50)]
        public int Id { get; set; }

        [Display(Name = "Full Name", Order = 1)]
        [Column(Width = 200)]
        public string? Name { get; set; }

        [Column(Format = "C2", NullDisplay = "$0.00")]
        public decimal? Amount { get; set; }

        [Column(IsHtml = true)]
        public string? HtmlContent { get; set; }

        private string Secret { get; set; } = "";
        protected internal string ProtectedInternal { get; set; } = "";
    }

    [Test]
    public void Property_ComplexScenario_AllAttributesCombine()
    {
        var nameProp = Properties<ComplexModel>.Items.First(p => p.Name == "Name");

        Assert.That(nameProp.DisplayName, Is.EqualTo("Full Name"));
        Assert.That(nameProp.Order, Is.EqualTo(1));
        Assert.That(nameProp.Width, Is.EqualTo(200));
        Assert.That(nameProp.Format, Is.Null);
        Assert.That(nameProp.IsHtml, Is.False);
    }

    [Test]
    public void Properties_ExcludesNonPublicProperties()
    {
        var items = Properties<ComplexModel>.Items;

        Assert.That(items, Has.None.Matches<Property<ComplexModel>>(p => p.Name == "Secret"));
        Assert.That(items, Has.None.Matches<Property<ComplexModel>>(p => p.Name == "ProtectedInternal"));
    }

    [Test]
    public void Property_Get_WorksWithNullableTypes()
    {
        var model = new ComplexModel
        {
            Amount = 123.45m
        };
        var prop = Properties<ComplexModel>.Items.First(p => p.Name == "Amount");

        Assert.That(prop.Get(model), Is.EqualTo(123.45m));
    }

    [Test]
    public void Property_Get_ReturnsNullForNullableWithNoValue()
    {
        var model = new ComplexModel
        {
            Amount = null
        };
        var prop = Properties<ComplexModel>.Items.First(p => p.Name == "Amount");

        Assert.That(prop.Get(model), Is.Null);
    }
}