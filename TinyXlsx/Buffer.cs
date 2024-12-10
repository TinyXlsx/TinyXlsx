using System.Globalization;
using System.Text;

namespace TinyXlsx;

public static class Buffer
{
    private static readonly byte[] buffer;
    private static readonly Encoder encoder;
    private static int bytesWritten;

    static Buffer()
    {
        buffer = new byte[1024 * 2];
        encoder = Encoding.UTF8.GetEncoder();
    }

    public static void Append(
        Stream stream,
        ReadOnlySpan<char> text)
    {
        while (text.Length > 0)
        {
            encoder.Convert(text, buffer.AsSpan(bytesWritten), false, out var charactersUsed, out var bytesUsed, out var isCompleted);

            bytesWritten += bytesUsed;

            if (bytesWritten + 4 > buffer.Length) Commit(stream);

            if (isCompleted) return;

            text = text[charactersUsed..];
        }
    }

    public static void Append(
        Stream stream,
        char character)
    {
        var singleChar = (Span<char>)[character];
        Append(stream, singleChar);
    }

    public static void Append(
        Stream stream,
        double value)
    {
        if (bytesWritten + Constants.MaximumDoubleLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    public static void Append(
        Stream stream,
        int value)
    {
        if (bytesWritten + Constants.MaximumIntegerLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    public static void Commit(Stream stream)
    {
        stream.Write(buffer, 0, bytesWritten);
        bytesWritten = 0;
    }
}
