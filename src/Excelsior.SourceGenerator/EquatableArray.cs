/// <summary>
/// An immutable array wrapper with structural equality for incremental generator caching.
/// </summary>
readonly struct EquatableArray<T>(ImmutableArray<T> array) :
    IEquatable<EquatableArray<T>>,
    IEnumerable<T>
    where T : IEquatable<T>
{
    public ImmutableArray<T> Array { get; } = array;

    public int Length => Array.Length;

    public T this[int index] => Array[index];

    public bool Equals(EquatableArray<T> other)
    {
        if (Array.Length != other.Array.Length)
        {
            return false;
        }

        for (var i = 0; i < Array.Length; i++)
        {
            if (!Array[i].Equals(other.Array[i]))
            {
                return false;
            }
        }

        return true;
    }

    public override bool Equals(object? obj) =>
        obj is EquatableArray<T> other && Equals(other);

    public override int GetHashCode()
    {
        var hash = 0;
        foreach (var item in Array)
        {
            hash = hash * 31 + item.GetHashCode();
        }

        return hash;
    }

    public ImmutableArray<T>.Enumerator GetEnumerator() => Array.GetEnumerator();

    IEnumerator<T> IEnumerable<T>.GetEnumerator() => ((IEnumerable<T>)Array).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable)Array).GetEnumerator();
}
