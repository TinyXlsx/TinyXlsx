using BenchmarkDotNet.Attributes;
using TinyXlsx;

[MemoryDiagnoser]
public partial class Benchmarks
{
    [Benchmark]
    public async Task TinyXlsx()
    {
        using var workbook = new Workbook();
        var stream = await workbook.BeginStreamAsync();
        using var worksheet = await workbook.BeginSheetAsync();

        for (var i = 0; i < 10_000; i++)
        {
            await worksheet.BeginRowAsync(i);
            await worksheet.WriteCellValueAsync(0, 123.456);
            await worksheet.WriteCellValueAsync(1, DateTime.Now);
            await worksheet.WriteCellValueAsync(2, "Text");
            await worksheet.WriteCellValueAsync(3, 123.456, "0.00");
            await worksheet.WriteCellValueAsync(4, 123.456, "0.00%");
            await worksheet.WriteCellValueAsync(5, 123.456, "0.00E+00");
            await worksheet.WriteCellValueAsync(6, 123.456, "$#,##0.00");
            await worksheet.WriteCellValueAsync(7, 123.456, "#,##0.00 [$USD]");
            await worksheet.EndRowAsync();
        }
        await workbook.EndSheetAsync();
        await workbook.EndStreamAsync();
    }
}
