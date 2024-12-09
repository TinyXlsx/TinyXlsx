namespace TinyXlsx;

internal static class ColumnKeyCache
{
    private static readonly Dictionary<int, string> cache;

    static ColumnKeyCache()
    {
        cache = new Dictionary<int, string>();
    }

    public static string GetKey(int columnIndex)
    {
        if (cache.TryGetValue(columnIndex, out var key))
        {
            return key;
        }

        // Maximum number of columns is 16,384 and thus XFD (3 characters).
        var needNewName = (Span<char>)stackalloc char[3];
        var i = 2;
        var needOtherNewName = columnIndex;

        while (needOtherNewName >= 0)
        {
            needNewName[i--] = (char)('A' + (needOtherNewName % 26));
            needOtherNewName = (needOtherNewName / 26) - 1;
        }
        
        var keyAsString = new string(needNewName.Slice(i + 1, 3 - (i + 1)));
        cache[columnIndex] = keyAsString;

        return keyAsString;
    }
}
