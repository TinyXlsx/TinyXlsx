using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace Benchmarks;

public partial class Benchmarks
{
    [Benchmark]
    public void ClosedXml()
    {
        using var workbook = new XLWorkbook();
        using var memoryStream = new MemoryStream();
        var sheet = workbook.Worksheets.Add();

        for (int i = 1; i < Records + 1; i++)
        {
            sheet.Cell(i, 1).Value = 123.456;
            sheet.Cell(i, 2).Value = DateTime.Now;
            sheet.Cell(i, 3).Value = "Text";
            
            var numberCell = sheet.Cell(i, 4);
            numberCell.Value = 123.456;
            numberCell.Style.NumberFormat.Format = "0.00";

            var percentageCell = sheet.Cell(i, 5);
            percentageCell.Value = 123.456;
            percentageCell.Style.NumberFormat.Format = "0.00%";

            var scientificCell = sheet.Cell(i, 6);
            scientificCell.Value = 123.456;
            scientificCell.Style.NumberFormat.Format = "0.00E+00";

            var currencyCell1 = sheet.Cell(i, 7);
            currencyCell1.Value = 123.456;
            currencyCell1.Style.NumberFormat.Format = "#,##0.00 [$USD]";

            var currencyCell2 = sheet.Cell(i, 8);
            currencyCell2.Value = 123.456;
            currencyCell2.Style.NumberFormat.Format = "$#,##0.00";
        }

        workbook.SaveAs(memoryStream);
    }
}
