using System.Buffers;
using System.Globalization;
using System.Text;

namespace TinyXlsx;

internal static class StreamExtensions
{
    private static readonly byte[] buffer = ArrayPool<byte>.Shared.Rent(2048);
    internal static void BufferPooledWrite(
    this Stream stream,
    string text)
    {
        var bytesWritten = Encoding.UTF8.GetBytes(text, buffer);
        stream.Write(buffer, 0, bytesWritten);
    }

    internal static void BufferPooledWrite(
        this Stream stream,
        char character)
    {
        stream.WriteByte((byte)character);
    }

    internal static void BufferPooledWrite(
        this Stream stream,
        double value)
    {
        value.TryFormat(buffer, out var bytesWritten, provider: CultureInfo.InvariantCulture);
        stream.Write(buffer, 0, bytesWritten);
    }

    internal static void BufferPooledWrite(
        this Stream stream,
        int value)
    {
        value.TryFormat(buffer, out var bytesWritten, provider: CultureInfo.InvariantCulture);
        stream.Write(buffer, 0, bytesWritten);
    }

    internal static void BufferPooledWrite(
        string text,
        ref int bytesWritten)
    {
        var written = Encoding.UTF8.GetBytes(text, 0, text.Length, buffer, bytesWritten);
        bytesWritten += written;
    }

    internal static void BufferPooledWrite(
        char character,
        ref int bytesWritten)
    {
        buffer[bytesWritten] = (byte)character;
        bytesWritten++;
    }

    internal static void BufferPooledWrite(
        double value,
        ref int bytesWritten)
    {
        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    internal static void BufferPooledWrite(
        int value,
        ref int bytesWritten)
    {
        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    internal static void WriteBufferedData(this Stream stream, int bytesWritten)
    {
        stream.Write(buffer, 0, bytesWritten);
    }
}
