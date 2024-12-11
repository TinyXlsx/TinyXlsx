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
            worksheet.WriteCellValue(123.456);
            worksheet.WriteCellValueAt(1, DateTime.Now);
            worksheet.WriteCellValueAt(2, DateTime.Now, "yyyy/MM/dd");
            worksheet.WriteCellValueAt(3, "Text");
            worksheet.WriteCellValueAt(4, 123.456, "0.00");
            worksheet.WriteCellValueAt(5, 123.456, "0.00%");
            worksheet.WriteCellValueAt(6, 123.456, "0.00E+00");
            worksheet.WriteCellValueAt(7, 123.456, "$#,##0.00");
            worksheet.WriteCellValueAt(8, 123.456, "#,##0.00 [$USD]");
        }
        workbook.Close();

        using var document = SpreadsheetDocument.Open(filePath, false);
        var validator = new OpenXmlValidator();
        var validationErrors = validator.Validate(document);

        Assert.IsTrue(!validationErrors.Any());
    }
}
