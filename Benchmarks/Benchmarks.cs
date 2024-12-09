using BenchmarkDotNet.Attributes;

namespace Benchmarks;

[MemoryDiagnoser]
public partial class Benchmarks
{
    //[Params(10, 1000, 100_000)]
    [Params(10_000)]
    public int Records { get; set; }
}
