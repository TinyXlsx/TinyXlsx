using BenchmarkDotNet.Attributes;
using TinyXlsx;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void TinyXlsx()
    {
        using var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        for (var i = 1; i <= Records; i++)
        {
            worksheet.BeginRow();
            worksheet.WriteCellValue(false);
            worksheet.WriteCellValue(123456);
            worksheet.WriteCellValue(123.456m);
            worksheet.WriteCellValue(123.456);
            worksheet.WriteCellValue(DateTime.Now);
            worksheet.WriteCellValue(DateTime.Now, "yyyy/MM/dd");
            worksheet.WriteCellValue("Text");
            worksheet.WriteCellValue(123.456, "0.00");
            worksheet.WriteCellValue(123.456, "0.00%");
            worksheet.WriteCellValue(123.456, "0.00E+00");
            worksheet.WriteCellValue(123.456, "$#,##0.00");
            worksheet.WriteCellValue(123.456, "#,##0.00 [$USD]");
        }
        var stream = workbook.Close();
    }
}
