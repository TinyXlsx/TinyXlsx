using BenchmarkDotNet.Attributes;
using FastExcel;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void FastExcel()
    {
        // FastExcel does not appear to support in-memory XLSX.
        using var stream = new MemoryStream();

        var worksheet = new Worksheet();
        var rows = new List<Row>();
        for (int i = 1; i <= Records; i++)
        {
            var cells = new List<Cell>
            {
                new (1, false),
                new (2, 123456),
                new (3, 123.456m),
                new (4, 123.456),
                new (5, DateTime.Now),
                new (6, DateTime.Now),
                new (7, "Text"),
                new (8, 123.456),
                new (9, 123.456),
                new (10, 123.456),
                new (11, 123.456),
                new (12, 123.456),
            };

            rows.Add(new Row(i, cells));
        }
        worksheet.Rows = rows;

        using FastExcel.FastExcel fastExcel = new(stream);
        fastExcel.Write(worksheet);
    }
}
