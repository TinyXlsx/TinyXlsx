using BenchmarkDotNet.Attributes;
using SharpCompress.Common;
using SharpCompress.Compressors.Deflate;
using SharpCompress.Writers;
using SharpCompress.Writers.Zip;
using System.Text;

namespace Benchmarks;

public partial class Compression
{
    [Benchmark]
    public void SharpCompress()
    {
        using var stream = new MemoryStream();
        using var zipWriter = (ZipWriter)WriterFactory.Open(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate) { DeflateCompressionLevel = CompressionLevel.Default });
        using var entryStream = zipWriter.WriteToStream("test", new ZipWriterEntryOptions());
        using var streamWriter = new StreamWriter(entryStream, Encoding.UTF8);
        
        for (var i = 0; i < Records; i++)
        {
            streamWriter.Write($"<c r=\"{i}\" s=\"2\" t=\"n\"><v>{i}</v></c>");
        }

        streamWriter.Dispose();
        entryStream.Dispose();
        zipWriter.Dispose();
    }
}
