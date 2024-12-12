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
        for (int rowNumber = 1; rowNumber <= Records; rowNumber++)
        {
            var cells = new List<Cell>
            {
                new (1, 123.456),
                new (2, DateTime.Now),
                new (3, "Text"),
                new (4, 123.456),
                new (5, 123.456),
                new (6, 123.456),
                new (7, 123.456),
                new (8, 123.456),
                new (9, 123.456),
            };

            rows.Add(new Row(rowNumber, cells));
        }
        worksheet.Rows = rows;

        using FastExcel.FastExcel fastExcel = new(stream);
        fastExcel.Write(worksheet);
    }
}
