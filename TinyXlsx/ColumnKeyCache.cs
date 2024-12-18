﻿namespace TinyXlsx;

/// <summary>
/// Provides efficent column index to column key conversion.
/// </summary>
public static class ColumnKeyCache
{
    private static readonly Dictionary<int, string> cache;

    static ColumnKeyCache()
    {
        cache = [];
    }

    /// <summary>
    /// Gets the column key, e.g. "A", "AB", for the specified column index.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based index of the column.
    /// </param>
    /// <returns>
    /// The column key.
    /// </returns>
    public static string GetKey(int columnIndex)
    {
        if (cache.TryGetValue(columnIndex, out var key))
        {
            return key;
        }

        // Maximum number of columns is 16,384 and thus XFD (3 characters).
        var keyBuffer = (Span<char>)stackalloc char[3];
        var i = 2;
        var remainingColumnIndex = columnIndex;

        while (remainingColumnIndex > 0)
        {
            var (quotient, remainder) = Math.DivRem(remainingColumnIndex - 1, 26);
            remainingColumnIndex = quotient;
            keyBuffer[i--] = (char)('A' + remainder);
        }

        var keyAsString = new string(keyBuffer[(i + 1)..3]);
        cache[columnIndex] = keyAsString;

        return keyAsString;
    }
}
