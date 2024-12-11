using System.Globalization;
using System.Text;

namespace TinyXlsx;

/// <summary>
/// Provides efficent buffering and writing of UTF-8 encoded text.
/// </summary>
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

    /// <summary>
    /// Appends a string of characters to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="text">
    /// The string of characters to append.
    /// </param>
    public static void Append(
        Stream stream,
        ReadOnlySpan<char> text)
    {
        const int MaximumUtf8BytesPerCharacter = 4;

        while (text.Length > 0)
        {
            encoder.Convert(text, buffer.AsSpan(bytesWritten), false, out var charactersUsed, out var bytesUsed, out var isCompleted);

            bytesWritten += bytesUsed;

            if (bytesWritten + MaximumUtf8BytesPerCharacter > buffer.Length) Commit(stream);

            if (isCompleted) return;

            text = text[charactersUsed..];
        }
    }

    /// <summary>
    /// Appends a single character to the internal buffer.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="character">
    /// The character to append.
    /// </param>
    public static void Append(
        Stream stream,
        char character)
    {
        var singleChar = (Span<char>)[character];
        Append(stream, singleChar);
    }

    /// <summary>
    /// Appends a <see cref="double"/> value to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="value">
    /// The <see cref="double"/> value to append.
    /// </param>
    public static void Append(
        Stream stream,
        double value)
    {
        if (bytesWritten + Constants.MaximumDoubleLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    /// <summary>
    /// Appends a <see cref="int"/> value to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="value">
    /// The <see cref="int"/> value to append.
    /// </param>
    public static void Append(
        Stream stream,
        int value)
    {
        if (bytesWritten + Constants.MaximumIntegerLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    /// <summary>
    /// Writes the contents of the internal buffer to the stream.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write the contents of the buffer to.
    /// </param>
    public static void Commit(Stream stream)
    {
        stream.Write(buffer, 0, bytesWritten);
        bytesWritten = 0;
    }
}
