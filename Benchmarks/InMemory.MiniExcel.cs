using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void MiniExcel()
    {
        using var stream = new MemoryStream();

        var configuration = new OpenXmlConfiguration
        {
            DynamicColumns = new DynamicExcelColumn[] {
                new("Column1"),
                new("Column2"),
                new("Column3"),
                new("Column4"),
                new("Column5") { Format="yyyy-MM-dd" },
                new("Column6") { Format="yyyy/MM/dd" },
                new("Column7"),
                new("Column8") { Format="0.00" },
                new("Column9") { Format="0.00%" },
                new("Column10") { Format="0.00E+00" },
                new("Column11") { Format="$#,##0.00" },
                new("Column12") { Format="#,##0.00 [$USD]" },
            },
        };

        var values = Enumerable
            .Range(1, Records)
            .Select(item =>
                new object[]
                {
                    new { Column1 = false },
                    new { Column2 = 123456 },
                    new { Column3 = 123.456m },
                    new { Column4 = 123.456 },
                    new { Column5 = DateTime.Now },
                    new { Column6 = DateTime.Now },
                    new { Column7 = "Text" },
                    new { Column8 = 123.456 },
                    new { Column9 = 123.456 },
                    new { Column10 = 123.456 },
                    new { Column11 = 123.456 },
                    new { Column12 = 123.456 },
                });

        stream.SaveAs(value: values, configuration: configuration, excelType: ExcelType.XLSX);
    }
}
