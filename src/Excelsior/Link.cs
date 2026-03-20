namespace Excelsior;

public record Link(string Text, string Url)
{
    public Link(string url) : this(url, url) { }
}
