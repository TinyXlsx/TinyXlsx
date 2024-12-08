﻿using BenchmarkDotNet.Attributes;
using TinyXlsx;

[MemoryDiagnoser]
public partial class Benchmarks
{
    [Benchmark]
    public async Task TinyXlsx()
    {
        using var workbook = new Workbook();
        var stream = await workbook.BeginStreamAsync();
        var worksheet = await workbook.BeginSheetAsync();

        for (var i = 0; i < 10_000; i++)
        {
            await worksheet.BeginRowAsync(i);
            worksheet.WriteCellValue(0, 123.456);
            worksheet.WriteCellValue(1, DateTime.Now);
            worksheet.WriteCellValue(2, "Text");
            worksheet.WriteCellValue(3, 123.456, "0.00");
            worksheet.WriteCellValue(4, 123.456, "0.00%");
            worksheet.WriteCellValue(5, 123.456, "0.00E+00");
            worksheet.WriteCellValue(6, 123.456, "$#,##0.00");
            worksheet.WriteCellValue(7, 123.456, "#,##0.00 [$USD]");
            await worksheet.EndRowAsync();
        }
        await workbook.EndSheetAsync();
        await workbook.EndStreamAsync();
    }
}
