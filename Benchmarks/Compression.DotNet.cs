using BenchmarkDotNet.Attributes;
using System.IO.Compression;
using System.Text;

namespace Benchmarks;

public partial class Compression
{
    [Benchmark]
    public void Dotnet()
    {
        using var stream = new MemoryStream();
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
        var entry = archive.CreateEntry("test");
        using var entryStream = entry.Open();
        using var streamWriter = new StreamWriter(entryStream, Encoding.UTF8);

        for (var i = 0; i < Records; i++)
        {
            streamWriter.Write($"<c r=\"{i}\" s=\"2\" t=\"n\"><v>{i}</v></c>");
        }
        streamWriter.Dispose();
        entryStream.Dispose();
        archive.Dispose();
    }
}
