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
        var utcDateStyle = workbook.CreateCellStyle();
        utcDateStyle.DataFormat = dataFormat.GetFormat("yyyy-MM-dd");
        var alternativeDateStyle = workbook.CreateCellStyle();
        alternativeDateStyle.DataFormat = dataFormat.GetFormat("yyyy/MM/dd");
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

            row.CreateCell(0).SetCellValue(false);

            row.CreateCell(1).SetCellValue(123456);

            row.CreateCell(2).SetCellValue((double)123.456m);

            row.CreateCell(3).SetCellValue(123.456);

            var utcDateCell = row.CreateCell(4);
            utcDateCell.SetCellValue(DateTime.Now);
            utcDateCell.CellStyle = utcDateStyle;

            var alternativeDateCell = row.CreateCell(5);
            alternativeDateCell.SetCellValue(DateTime.Now);
            alternativeDateCell.CellStyle = alternativeDateStyle;

            row.CreateCell(6).SetCellValue("Text");

            var numberCell = row.CreateCell(7);
            numberCell.SetCellValue(123.456);
            numberCell.CellStyle = numberStyle;

            var percentageCell = row.CreateCell(8);
            percentageCell.SetCellValue(123.456);
            percentageCell.CellStyle = percentageStyle;

            var scientificCell = row.CreateCell(9);
            scientificCell.SetCellValue(123.456);
            scientificCell.CellStyle = scientificStyle;

            var currencyCell1 = row.CreateCell(10);
            currencyCell1.SetCellValue(123.456);
            currencyCell1.CellStyle = currencyStyle1;

            var currencyCell2 = row.CreateCell(11);
            currencyCell2.SetCellValue(123.456);
            currencyCell2.CellStyle = currencyStyle2;
        }

        workbook.Write(memoryStream);
    }
}
