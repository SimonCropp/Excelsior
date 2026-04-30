namespace Excelsior;

static class ProtectionPasswordHasher
{
    public static HexBinaryValue Hash(string password)
    {
        if (password.Length == 0)
        {
            return new("0000");
        }

        ushort hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF));
            hash ^= password[i];
        }

        hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF));
        hash ^= 0x8000 | ('N' << 8) | 'K';
        hash ^= (ushort)password.Length;
        return new(hash.ToString("X4"));
    }
}
