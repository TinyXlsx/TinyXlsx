using BenchmarkDotNet.Attributes;

namespace Benchmarks;

[MemoryDiagnoser]
public partial class Compression
{
    [Params(100, 10_000, 1_000_000)]
    public int Records { get; set; }
}
