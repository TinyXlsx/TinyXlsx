using System.Buffers;
using System.Globalization;
using System.Text;

namespace TinyXlsx;

internal static class StreamExtensions
{
    //private static readonly Encoder Utf8Encoder = Encoding.UTF8.GetEncoder();
    private static readonly byte[] buffer = ArrayPool<byte>.Shared.Rent(2048);

    internal static void BufferPooledWrite(
        this Stream stream,
        string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        //var buffer = ArrayPool<byte>.Shared.Rent(Encoding.UTF8.GetMaxByteCount(text.Length));
        //var buffer = ArrayPool<byte>.Shared.Rent(Utf8Encoder.GetByteCount(text, false));
        //try
        //{
            var bytesWritten = Encoding.UTF8.GetBytes(text, buffer);
            //var bytesWritten = Utf8Encoder.GetBytes(text, buffer, true);
            //await stream.WriteAsync(buffer.AsMemory(0, bytesWritten));
            stream.Write(buffer, 0, bytesWritten);
        //}
        //finally
        //{
        //    ArrayPool<byte>.Shared.Return(buffer);
        //}
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
        //var buffer = ArrayPool<byte>.Shared.Rent(32);
        value.TryFormat(buffer, out var bytesWritten, provider: CultureInfo.InvariantCulture);
        //try
        //{
            //await stream.WriteAsync(buffer.AsMemory(0, bytesWritten));
            stream.Write(buffer, 0, bytesWritten);
        //}
        //finally
        //{
        //    ArrayPool<byte>.Shared.Return(buffer);
        //}
    }

    internal static void BufferPooledWrite(
        this Stream stream,
        int value)
    {
        //var buffer = ArrayPool<byte>.Shared.Rent(11);
        value.TryFormat(buffer, out var bytesWritten, provider: CultureInfo.InvariantCulture);
        //try
        //{
            //await stream.WriteAsync(buffer.AsMemory(0, bytesWritten));
            stream.Write(buffer, 0, bytesWritten);
        //}
        //finally
        //{
        //    ArrayPool<byte>.Shared.Return(buffer);
        //}
    }
}
