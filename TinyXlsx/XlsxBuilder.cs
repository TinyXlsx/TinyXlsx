using System.Globalization;
using System.Text;

namespace TinyXlsx;

/// <summary>
/// Provides efficent buffering and writing of UTF-8 encoded XLSX.
/// </summary>
public class XlsxBuilder
{
    private readonly byte[] buffer;
    private static readonly Encoder encoder;
    private int bytesWritten;

    static XlsxBuilder()
    {
        encoder = Encoding.UTF8.GetEncoder();
    }

    public XlsxBuilder()
    {
        buffer = new byte[1024 * 8];
    }

    /// <summary>
    /// Appends a <see cref="bool"/> value to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="value">
    /// The <see cref="bool"/> value to append.
    /// </param>
    public void Append(
        Stream stream,
        bool value)
    {
        if (bytesWritten == buffer.Length) Commit(stream);

        if (value)
        {
            buffer[bytesWritten++] = 0x31;
        }
        else
        {
            buffer[bytesWritten++] = 0x30;
        }
    }

    /// <summary>
    /// Appends a <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="bytes">
    /// The <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> to append.
    /// </param>
    public void Append(
        Stream stream,
        ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length + bytesWritten > buffer.Length) Commit(stream);

        bytes.CopyTo(buffer.AsSpan(bytesWritten));
        bytesWritten += bytes.Length;
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
    public void Append(
        Stream stream,
        ReadOnlySpan<char> text)
    {
        const int MaximumUtf8BytesPerCharacter = 4;

        while (text.Length > 0)
        {
            if (bytesWritten + MaximumUtf8BytesPerCharacter > buffer.Length) Commit(stream);

            var lessThanIndex = text.IndexOf('<');
            var ampersandIndex = text.IndexOf('&');

            if (lessThanIndex >= 0 && (lessThanIndex < ampersandIndex || ampersandIndex == -1))
            {
                encoder.Convert(text[..lessThanIndex], buffer.AsSpan(bytesWritten), false, out _, out var bytesUsed, out _);
                bytesWritten += bytesUsed;

                if (bytesWritten + 4 > buffer.Length) Commit(stream);

                buffer[bytesWritten++] = (byte)'&';
                buffer[bytesWritten++] = (byte)'l';
                buffer[bytesWritten++] = (byte)'t';
                buffer[bytesWritten++] = (byte)';';
                text = text[(lessThanIndex + 1)..];
            }
            else if (ampersandIndex >= 0 && (ampersandIndex < lessThanIndex || lessThanIndex == -1))
            {
                encoder.Convert(text[..ampersandIndex], buffer.AsSpan(bytesWritten), false, out _, out var bytesUsed, out _);
                bytesWritten += bytesUsed;

                if (bytesWritten + 5 > buffer.Length) Commit(stream);

                buffer[bytesWritten++] = (byte)'&';
                buffer[bytesWritten++] = (byte)'a';
                buffer[bytesWritten++] = (byte)'m';
                buffer[bytesWritten++] = (byte)'p';
                buffer[bytesWritten++] = (byte)';';
                text = text[(ampersandIndex + 1)..];
            }
            else
            {
                encoder.Convert(text, buffer.AsSpan(bytesWritten), false, out var charactersUsed, out var bytesUsed, out _);
                text = text[charactersUsed..];
                bytesWritten += bytesUsed;
            }
        }
    }

    /// <summary>
    /// Appends a <see cref="decimal"/> value to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="value">
    /// The <see cref="decimal"/> value to append.
    /// </param>
    public void Append(
        Stream stream,
        decimal value)
    {
        if (bytesWritten + Constants.MaximumDecimalLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;
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
    public void Append(
        Stream stream,
        double value)
    {
        if (bytesWritten + Constants.MaximumDoubleLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;
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
    public void Append(
        Stream stream,
        int value)
    {
        if (bytesWritten + Constants.MaximumIntegerLength > buffer.Length) Commit(stream);

        value.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;
    }

    /// <summary>
    /// Appends the specified column index in column key form to the internal buffer and writes to the stream if the buffer size will be exceeded.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write to when the buffer is full.
    /// </param>
    /// <param name="columnIndex">
    /// The one-based index of the column to convert to a key and append.
    /// </param>
    public void AppendColumnKey(
        Stream stream,
        int columnIndex)
    {
        if (bytesWritten + 3 > buffer.Length) Commit(stream);

        if (columnIndex <= 26)
        {
            buffer[bytesWritten++] = (byte)(64 + columnIndex);
            return;
        }

        Append(stream, ColumnKeyCache.GetKey(columnIndex));
    }

    /// <summary>
    /// Writes the contents of the internal buffer to the stream.
    /// </summary>
    /// <param name="stream">
    /// The target <see cref="Stream"/> to write the contents of the buffer to.
    /// </param>
    public void Commit(Stream stream)
    {
        stream.Write(buffer, 0, bytesWritten);
        bytesWritten = 0;
    }
}
