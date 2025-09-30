public class NullableTargets
{
    public required int? Number { get; init; }
    public required string? String { get; init; }
    public required DateTime? DateTime { get; init; }
    public required AnEnum? Enum { get; init; }
    public required bool? Bool { get; init; }

    public static IReadOnlyList<NullableTargets> Data { get; } =
    [
        new()
        {
            Number = null,
            String = null,
            DateTime = null,
            Enum = null,
            Bool = null
        },
        new()
        {
            Number = 1,
            String = "value",
            DateTime = new DateTime(2020, 1, 1),
            Enum = AnEnum.Value,
            Bool = true
        },
    ];
}