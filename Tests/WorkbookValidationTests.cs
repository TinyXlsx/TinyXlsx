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

        for (var i = 0; i < 1; i++)
        {
            worksheet.BeginRowAt(i);
            worksheet.WriteCellValue(true);
            worksheet.WriteCellValue(0.1m);
            worksheet.WriteCellValue(123.456);
            worksheet.WriteCellValue(DateTime.Now);
            worksheet.WriteCellValue(DateTime.Now, "yyyy/MM/dd");
            worksheet.WriteCellValue("Text");
            worksheet.WriteCellValue(123.456, "0.00");
            worksheet.WriteCellValue(123.456, "0.00%");
            worksheet.WriteCellValue(123.456, "0.00E+00");
            worksheet.WriteCellValue(123.456, "$#,##0.00");
            worksheet.WriteCellValue(123.456, "#,##0.00 [$USD]");

            worksheet.WriteCellFormula("=SUM(B1:C1)");
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
