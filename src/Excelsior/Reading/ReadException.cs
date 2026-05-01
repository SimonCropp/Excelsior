namespace Excelsior;

public sealed class ReadException(IReadOnlyList<ReadError> errors) :
    Exception(BuildMessage(errors))
{
    public IReadOnlyList<ReadError> Errors { get; } = errors;

    static string BuildMessage(IReadOnlyList<ReadError> errors)
    {
        if (errors.Count == 0)
        {
            return "Read failed.";
        }

        var builder = new StringBuilder();
        builder.Append("Read failed with ");
        builder.Append(errors.Count);
        builder.AppendLine(errors.Count == 1 ? " error:" : " errors:");

        var max = Math.Min(errors.Count, 10);
        for (var i = 0; i < max; i++)
        {
            builder.Append(" - ");
            builder.AppendLine(errors[i].ToString());
        }

        if (errors.Count > max)
        {
            builder.Append(" - (");
            builder.Append(errors.Count - max);
            builder.AppendLine(" more)");
        }

        return builder.ToString();
    }
}
