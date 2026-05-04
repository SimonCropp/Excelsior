static class ModelActivator<T>
{
    static Func<IReadOnlyDictionary<string, object?>, T>? generatedFactory;
    static ConstructorInfo? parameterlessCtor;
    static ConstructorInfo? matchingCtor;
    static string[] ctorParamNames;
    static HashSet<string> ctorParamSet;
    static Dictionary<string, Action<T, object?>> setters;

    static ModelActivator()
    {
        generatedFactory = GeneratedActivators.TryGet<T>();
        if (generatedFactory != null)
        {
            ctorParamNames = [];
            ctorParamSet = new(StringComparer.Ordinal);
            setters = new();
            return;
        }

        var type = typeof(T);

        parameterlessCtor = type.GetConstructor(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: Type.EmptyTypes,
            modifiers: null);

        setters = BuildSetters(type);

        if (parameterlessCtor != null)
        {
            matchingCtor = null;
            ctorParamNames = [];
            ctorParamSet = new(StringComparer.Ordinal);
            return;
        }

        matchingCtor = type
            .GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .OrderByDescending(_ => _.GetParameters().Length)
            .FirstOrDefault();

        ctorParamNames = matchingCtor
            ?.GetParameters()
            .Select(_ => _.Name ?? "")
            .ToArray() ?? [];
        ctorParamSet = new(ctorParamNames, StringComparer.Ordinal);
    }

    static Dictionary<string, Action<T, object?>> BuildSetters(Type type)
    {
        var result = new Dictionary<string, Action<T, object?>>(StringComparer.Ordinal);
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanWrite)
            {
                continue;
            }

            var setter = property.SetMethod;
            if (setter == null)
            {
                continue;
            }

            result[property.Name] = BuildSetter(property, setter);
        }

        return result;
    }

    static Action<T, object?> BuildSetter(PropertyInfo property, MethodInfo setter)
    {
        var target = Expression.Parameter(typeof(T), "target");
        var value = Expression.Parameter(typeof(object), "value");
        var converted = Expression.Convert(value, property.PropertyType);
        var call = Expression.Call(target, setter, converted);
        return Expression.Lambda<Action<T, object?>>(call, target, value).Compile();
    }

    public static T Create(IReadOnlyDictionary<string, object?> values)
    {
        if (generatedFactory != null)
        {
            return generatedFactory(values);
        }

        T instance;
        if (parameterlessCtor != null)
        {
            instance = (T)parameterlessCtor.Invoke(null);
        }
        else if (matchingCtor != null)
        {
            var args = new object?[ctorParamNames.Length];
            for (var i = 0; i < ctorParamNames.Length; i++)
            {
                values.TryGetValue(ctorParamNames[i], out args[i]);
            }

            instance = (T)matchingCtor.Invoke(args);
        }
        else
        {
            throw new($"Type {typeof(T).Name} has no usable constructor for deserialization. Provide a parameterless constructor or a constructor whose parameter names match the column property names.");
        }

        foreach (var (name, value) in values)
        {
            if (parameterlessCtor == null &&
                ctorParamSet.Contains(name))
            {
                continue;
            }

            if (setters.TryGetValue(name, out var setter))
            {
                setter(instance, value);
            }
        }

        return instance;
    }
}
