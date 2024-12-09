using System.Buffers;
using System.Globalization;
using System.Text;

namespace TinyXlsx;

public class Buffer : IDisposable
{
    private readonly byte[] buffer;
    private readonly Encoder encoder;
    public int bytesWritten;
    private bool disposedValue;
    private readonly Stream stream;

    public Buffer(Stream stream)
    {
        buffer = ArrayPool<byte>.Shared.Rent(1024 * 2);
        encoder = Encoding.UTF8.GetEncoder();
        this.stream = stream;
    }

    public void Append(ReadOnlySpan<char> text)
    {
        while (text.Length > 0)
        {
            encoder.Convert(text, buffer.AsSpan(bytesWritten), false, out var charactersUsed, out var bytesUsed, out var isCompleted);

            bytesWritten += bytesUsed;

            if (bytesWritten + 4 > buffer.Length) Commit();

            if (isCompleted) return;

            text = text[charactersUsed..];
        }
    }

    public void Append(char character)
    {
        if (bytesWritten >= buffer.Length) Commit();

        buffer[bytesWritten] = (byte)character;
        bytesWritten++;
    }

    public void Append(double value)
    {
        if (bytesWritten + 64 > buffer.Length) Commit();

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    public void Append(int value)
    {
        if (bytesWritten + 64 > buffer.Length) Commit();

        value.TryFormat(buffer.AsSpan(bytesWritten), out var written, provider: CultureInfo.InvariantCulture);
        bytesWritten += written;
    }

    public void Commit()
    {
        stream.Write(buffer, 0, bytesWritten);
        bytesWritten = 0;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }

            disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
