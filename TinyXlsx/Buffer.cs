using System.Buffers;
using System.Globalization;
using System.Text;

namespace TinyXlsx;

internal static class Buffer
{
    private static readonly byte[] buffer = ArrayPool<byte>.Shared.Rent(2048);
    private static int bytesWritten;

    internal static void Append(string text)
    {
        var written = Encoding.UTF8.GetBytes(text, 0, text.Length, buffer, bytesWritten);
        bytesWritten += written;
    }

    internal static void Append(char character)
    {
        buffer[bytesWritten] = (byte)character;
        bytesWritten++;
    }

    internal static void Append(double value)
    {
        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    internal static void Append(int value)
    {
        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    internal static void Commit(Stream stream)
    {
        stream.Write(buffer, 0, bytesWritten);
        bytesWritten = 0;
    }

    internal static byte[] Get()
    {
        return buffer;
    }
}
