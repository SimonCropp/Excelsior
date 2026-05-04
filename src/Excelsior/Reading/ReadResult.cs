namespace Excelsior;

public sealed class ReadResult(IReadOnlyList<ReadError> errors)
{
    public IReadOnlyList<ReadError> Errors { get; } = errors;

    public bool Succeeded => Errors.Count == 0;

    public static implicit operator bool(ReadResult result) =>
        result.Succeeded;

    public static implicit operator ReadError[](ReadResult result) =>
        [..result.Errors];

    public static implicit operator List<ReadError>(ReadResult result) =>
        [..result.Errors];
}
