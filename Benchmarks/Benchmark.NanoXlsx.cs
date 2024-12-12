using BenchmarkDotNet.Attributes;

namespace Benchmarks;

public partial class Benchmarks
{
    [Benchmark]
    public void NanoXlsx()
    {
        // NanoXLSX does not support in-memory XLSX.
        throw new NotSupportedException();
    }
}
