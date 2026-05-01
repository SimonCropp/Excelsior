namespace Excelsior;

public sealed class ReadException(IReadOnlyList<ReadError> errors) :
    Exception(BuildMessage(errors))
{
    public IReadOnlyList<ReadError> Errors { get; } = errors;

    static string BuildMessage(IReadOnlyList<ReadError> errors)
    {
        var count = errors.Count;
        if (count == 0)
        {
            return "Read failed.";
        }

        var builder = new StringBuilder();
        builder.Append("Read failed with ");
        builder.Append(count);
        builder.AppendLine(count == 1 ? " error:" : " errors:");

        foreach (var error in errors)
        {
            builder.Append(" - ");
            builder.AppendLine(error.ToString());
        }

        return builder.ToString();
    }
}
