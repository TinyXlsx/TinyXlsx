using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void MiniExcel()
    {
        using var stream = new MemoryStream();

        var values = Enumerable
            .Range(1, Records)
            .Select(item =>
                new Dto
                {
                    Column1 = false,
                    Column2 = 123456,
                    Column3 = 123.456m,
                    Column4 = 123.456,
                    Column5 = DateTime.Now,
                    Column6 = DateTime.Now,
                    Column7 = "Text",
                    Column8 = 123.456,
                    Column9 = 123.456,
                    Column10 = 123.456,
                    Column11 = 123.456,
                    Column12 = 123.456,
                });

        stream.SaveAs(value: values, excelType: ExcelType.XLSX);
    }

    public class Dto
    {
        public bool Column1 { get; set; }

        public int Column2 { get; set; }

        public decimal Column3 { get; set; }

        public double Column4 { get; set; }

        [ExcelFormat("yyyy-MM-dd")]
        public DateTime Column5 { get; set; }

        [ExcelFormat("yyyy/MM/dd")]
        public DateTime Column6 { get; set; }

        public string Column7 { get; set; }

        [ExcelFormat("0.00")]
        public double Column8 { get; set; }

        [ExcelFormat("0.00%")]
        public double Column9 { get; set; }

        [ExcelFormat("0.00E+00")]
        public double Column10 { get; set; }

        [ExcelFormat("$#,##0.00")]
        public double Column11 { get; set; }

        [ExcelFormat("#,##0.00 [$USD]")]
        public double Column12 { get; set; }
    }
}
