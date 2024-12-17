using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using TinyXlsx;

namespace Tests;

[TestClass]
public class WorkbookValidationTests
{
    [TestMethod]
    public void GeneratedXlsxFileShouldHaveNoValidationErrors()
    {
        var filePath = "test.xlsx";
        using var workbook = new Workbook(filePath);
        var worksheet = workbook.BeginSheet();

        var i = 1;
        for (; i <= 10; i++)
        {
            worksheet.BeginRow();
            worksheet.WriteCellValue(true);
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

            worksheet.WriteCellFormula($"=SUM(H{i}:L{i})");
        }
        workbook.Close();

        using var document = SpreadsheetDocument.Open(filePath, false);
        var validator = new OpenXmlValidator();
        var validationErrors = validator.Validate(document);

        foreach (var validationError in validationErrors)
        {
            Console.WriteLine(validationError.Description);
        }

        Assert.IsTrue(!validationErrors.Any());
    }
}
