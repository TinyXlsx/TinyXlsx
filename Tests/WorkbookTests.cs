using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using TinyXlsx;

namespace Tests;

[TestClass]
public class WorkbookTests
{
    [TestMethod]
    public void GeneratedXlsxFileShouldHaveNoValidationErrors()
    {
        var filePath = "test.xlsx";
        using var workbook = new Workbook(filePath);
        var worksheet = workbook.BeginSheet();

        var i = 1;
        for (; i <= 10_000; i++)
        {
            worksheet
                .BeginRow()
                .WriteCellValue(true)
                .WriteCellValue(123456)
                .WriteCellValue(123.456m)
                .WriteCellValue(123.456)
                .WriteCellValue(DateTime.Now)
                .WriteCellValue(DateTime.Now, "yyyy/MM/dd")
                .WriteCellValue("Text")
                .WriteCellValue(123.456, "0.00")
                .WriteCellValue(123.456, "0.00%")
                .WriteCellValue(123.456, "0.00E+00")
                .WriteCellValue(123.456, "$#,##0.00")
                .WriteCellValue(123.456, "#,##0.00 [$USD]")
                .WriteCellFormula($"=SUM(H{i}:L{i})");
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
