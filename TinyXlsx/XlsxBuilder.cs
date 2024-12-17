using System.Globalization;
using System.Runtime.CompilerServices;
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
            encoder.Convert(text, buffer.AsSpan(bytesWritten), false, out var charactersUsed, out var bytesUsed, out var isCompleted);

            bytesWritten += bytesUsed;

            if (bytesWritten + MaximumUtf8BytesPerCharacter > buffer.Length) Commit(stream);

            if (isCompleted) return;

            text = text[charactersUsed..];
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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void AppendCellValueAt(
        Stream stream,
        int columnIndex,
        int rowIndex,
        bool value)
    {
        if (bytesWritten + 50 > buffer.Length) Commit(stream);

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x72;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;

        if (columnIndex <= 26)
        {
            buffer[bytesWritten++] = (byte)(64 + columnIndex);
        }
        else
        {
            Append(stream, ColumnKeyCache.GetKey(columnIndex));
        }

        rowIndex.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x74;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x62;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;

        if (value)
        {
            buffer[bytesWritten++] = 0x31;
        }
        else
        {
            buffer[bytesWritten++] = 0x30;
        }

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x3E;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void AppendCellValueAt(
        Stream stream,
        int columnIndex,
        int rowIndex,
        double value)
    {
        if (bytesWritten + 50 > buffer.Length) Commit(stream);

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x72;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;

        if (columnIndex <= 26)
        {
            buffer[bytesWritten++] = (byte)(64 + columnIndex);
        }
        else
        {
            Append(stream, ColumnKeyCache.GetKey(columnIndex));
        }

        rowIndex.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x74;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x6E;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;

        value.TryFormat(buffer.AsSpan(bytesWritten), out bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x3E;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void AppendCellValueAt(
        Stream stream,
        int columnIndex,
        int rowIndex,
        int styleIndex,
        double value)
    {
        if (bytesWritten + 75 > buffer.Length) Commit(stream);

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x72;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;

        if (columnIndex <= 26)
        {
            buffer[bytesWritten++] = (byte)(64 + columnIndex);
        }
        else
        {
            Append(stream, ColumnKeyCache.GetKey(columnIndex));
        }

        rowIndex.TryFormat(buffer.AsSpan(bytesWritten), out var bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x73;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;

        styleIndex.TryFormat(buffer.AsSpan(bytesWritten), out bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x20;
        buffer[bytesWritten++] = 0x74;
        buffer[bytesWritten++] = 0x3D;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x6E;
        buffer[bytesWritten++] = 0x22;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;

        value.TryFormat(buffer.AsSpan(bytesWritten), out bytesUsed, provider: CultureInfo.InvariantCulture);
        bytesWritten += bytesUsed;

        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x76;
        buffer[bytesWritten++] = 0x3E;
        buffer[bytesWritten++] = 0x3C;
        buffer[bytesWritten++] = 0x2F;
        buffer[bytesWritten++] = 0x63;
        buffer[bytesWritten++] = 0x3E;
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
