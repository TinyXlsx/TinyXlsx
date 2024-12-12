using BenchmarkDotNet.Attributes;
using NPOI.XSSF.Streaming;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void Npoi()
    {
        using var workbook = new SXSSFWorkbook();
        using var memoryStream = new MemoryStream();
        var sheet = workbook.CreateSheet();

        var dataFormat = workbook.CreateDataFormat();
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = dataFormat.GetFormat("yyyy-MM-dd");
        var numberStyle = workbook.CreateCellStyle();
        numberStyle.DataFormat = dataFormat.GetFormat("0.00");
        var percentageStyle = workbook.CreateCellStyle();
        percentageStyle.DataFormat = dataFormat.GetFormat("0.00%");
        var scientificStyle = workbook.CreateCellStyle();
        scientificStyle.DataFormat = dataFormat.GetFormat("0.00E+00");
        var currencyStyle1 = workbook.CreateCellStyle();
        currencyStyle1.DataFormat = dataFormat.GetFormat("#,##0.00 [$USD]");
        var currencyStyle2 = workbook.CreateCellStyle();
        currencyStyle2.DataFormat = dataFormat.GetFormat("$#,##0.00");

        for (int i = 0; i < Records; i++)
        {
            var row = sheet.CreateRow(i);

            row.CreateCell(0).SetCellValue(123.456);

            var dateCell = row.CreateCell(1);
            dateCell.SetCellValue(DateTime.Now);
            dateCell.CellStyle = dateStyle;

            row.CreateCell(2).SetCellValue("Text");

            var numberCell = row.CreateCell(3);
            numberCell.SetCellValue(123.456);
            numberCell.CellStyle = numberStyle;

            var percentageCell = row.CreateCell(4);
            percentageCell.SetCellValue(123.456);
            percentageCell.CellStyle = percentageStyle;

            var scientificCell = row.CreateCell(5);
            scientificCell.SetCellValue(123.456);
            scientificCell.CellStyle = scientificStyle;

            var currencyCell1 = row.CreateCell(6);
            currencyCell1.SetCellValue(123.456);
            currencyCell1.CellStyle = currencyStyle1;

            var currencyCell2 = row.CreateCell(7);
            currencyCell2.SetCellValue(123.456);
            currencyCell2.CellStyle = currencyStyle2;
        }

        workbook.Write(memoryStream);
    }
}
