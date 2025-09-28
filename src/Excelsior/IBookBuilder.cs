namespace Excelsior;

public interface IBookBuilder
{
    public Task ToStream(Stream stream, Cancel cancel = default);

    public async Task ToFile(string path, Cancel cancel = default)
    {
        await using var stream = File.Create(path);
        await ToStream(stream, cancel);
    }
}