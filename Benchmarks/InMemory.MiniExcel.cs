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
                new DynamicExcelColumn("Column1"),
                new DynamicExcelColumn("Column2") { Format="yyyy-MM-dd" },
                new DynamicExcelColumn("Column3"),
                new DynamicExcelColumn("Column4") { Format="0.00" },
                new DynamicExcelColumn("Column5") { Format="0.00%" },
                new DynamicExcelColumn("Column6") { Format="0.00E+00" },
                new DynamicExcelColumn("Column7") { Format="$#,##0.00" },
                new DynamicExcelColumn("Column8") { Format="#,##0.00 [$USD]" },
            },
        };

        var values = Enumerable
            .Range(1, Records)
            .Select(item =>
                new object[]
                {
                    new { Column1 = 123.456 },
                    new { Column2 = DateTime.Now },
                    new { Column3 = "Text" },
                    new { Column4 = 123.456 },
                    new { Column5 = 123.456 },
                    new { Column6 = 123.456 },
                    new { Column7 = 123.456 },
                    new { Column8 = 123.456 },
                });

        stream.SaveAs(value: values, configuration: configuration, excelType: ExcelType.XLSX);
    }
}
