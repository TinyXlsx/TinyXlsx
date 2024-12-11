using TinyXlsx;

namespace Tests;

[TestClass]
public class WorksheetConstraintTests
{
    [TestMethod]
    public void BeginRowAtWithLowerIndexThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        worksheet.BeginRowAt(1);

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.BeginRowAt(0));
    }

    [TestMethod]
    public void BeginRowAtWithSameIndexThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        worksheet.BeginRowAt(1);

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.BeginRowAt(1));
    }

    [TestMethod]
    public void WriteCellValueBeforeBeginRowThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.WriteCellValue(123.456));
    }

    [TestMethod]
    public void WriteCellValueAtWithSameIndexThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        worksheet.BeginRow();
        worksheet.WriteCellValueAt(0, 123.456);

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.WriteCellValueAt(0, 123.456));
    }

    [TestMethod]
    public void BeginSheetWithSameNameThrowsException()
    {
        const string name = "SheetName";
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet(name);

        Assert.ThrowsException<InvalidOperationException>(() => workbook.BeginSheet(name));
    }

    [TestMethod]
    public void BeginRowAtWithHigherValueThanMaximumRowsThrowsException()
    {
        var rowIndex = Constants.MaximumRows + 1;
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.BeginRowAt(rowIndex));
    }

    [TestMethod]
    public void WriteCellValueAtAtWithHigherValueThanMaximumColumnsThrowsException()
    {
        var columnIndex = Constants.MaximumColumns + 1;
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();
        worksheet.BeginRow();
        Assert.ThrowsException<InvalidOperationException>(() => worksheet.WriteCellValueAt(columnIndex, "test"));
    }

    [TestMethod]
    public void WriteCellValueWithLowerDateTimeValueThanMinimumDateThrowsException()
    {
        var date = Constants.MinimumDate.AddDays(-1);
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();
        worksheet.BeginRow();
        Assert.ThrowsException<NotSupportedException>(() => worksheet.WriteCellValue(date));
    }

    [TestMethod]
    public void WriteCellValueWithMinimumDateSucceeds()
    {
        var date = Constants.MinimumDate;
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();
        worksheet.BeginRow();
        worksheet.WriteCellValue(date);
    }

    [TestMethod]
    public void WriteCellValueWithMaximumDateSucceeds()
    {
        var date = DateTime.MaxValue;
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();
        worksheet.BeginRow();
        worksheet.WriteCellValue(date);
    }

    [TestMethod]
    public void WriteCellValueWithStringLongerThanMaximumCharactersPerCellThrowsException()
    {
        var workbook = new Workbook();
        var longstring = new string('a', Constants.MaximumCharactersPerCell + 1);

        var worksheet = workbook.BeginSheet();
        worksheet.BeginRow();

        Assert.ThrowsException<NotSupportedException>(() => worksheet.WriteCellValue(longstring));
    }

    [TestMethod]
    public void ExceedMaximumNumberOfStylesThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.BeginSheet();
        for (var i = 0; i < Constants.MaximumStyles; i++)
        {
            worksheet.BeginRow();
            worksheet.WriteCellValue(123.456, i.ToString());
        }

        Assert.ThrowsException<NotSupportedException>(() => worksheet.WriteCellValue(123.456, "0.00"));
    }
}
