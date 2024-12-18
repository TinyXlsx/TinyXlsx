using TinyXlsx;

namespace Tests;

[TestClass]
public class WorksheetTests
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
        worksheet.WriteCellValueAt(1, 123.456);

        Assert.ThrowsException<InvalidOperationException>(() => worksheet.WriteCellValueAt(1, 123.456));
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

    [DynamicData(nameof(GetWriteCellValueWithNullSkipsColumnData), DynamicDataSourceType.Method)]
    [TestMethod]
    public void WriteCellValueWithNullSkipsColumn(
        object value,
        string expected)
    {
        var memoryStream = new MemoryStream();
        var xlsxBuilder = new XlsxBuilder();
        var stylesheet = new Stylesheet();
        var worksheet = new Worksheet(
            xlsxBuilder,
            memoryStream,
            stylesheet,
            1,
            "Sheet1",
            "rId3");

        worksheet.BeginRow();
        switch (value)
        {
            case bool boolValue:
                worksheet.WriteCellValue((bool?)null);
                worksheet.WriteCellValue(boolValue);
                break;
            case DateTime dateTimeValue:
                worksheet.WriteCellValue((DateTime?)null);
                worksheet.WriteCellValue(dateTimeValue);
                break;
            case decimal decimalValue:
                worksheet.WriteCellValue((decimal?)null);
                worksheet.WriteCellValue(decimalValue);
                break;
            case double doubleValue:
                worksheet.WriteCellValue((double?)null);
                worksheet.WriteCellValue(doubleValue);
                break;
            case string stringValue:
                worksheet.WriteCellValue((string?)null);
                worksheet.WriteCellValue(stringValue);
                break;
        }
        xlsxBuilder.Commit(memoryStream);

        memoryStream.Position = 0;
        using var reader = new StreamReader(memoryStream);
        var result = reader.ReadToEnd();
        Assert.AreEqual(expected, result);
    }

    private static IEnumerable<object[]> GetWriteCellValueWithNullSkipsColumnData()
    {
        yield return new object[]
        {
            true,
            "<row r=\"1\"><c r=\"B1\" t=\"b\"><v>1</v></c>",
        };

        yield return new object[]
        {
            new DateTime(2024, 1, 1),
            "<row r=\"1\"><c r=\"B1\" s=\"1\" t=\"n\"><v>45292</v></c>",
        };

        yield return new object[]
        {
            123.456m,
            "<row r=\"1\"><c r=\"B1\" t=\"n\"><v>123.456</v></c>",
        };

        yield return new object[]
        {
            123.456,
            "<row r=\"1\"><c r=\"B1\" t=\"n\"><v>123.456</v></c>",
        };

        yield return new object[]
        {
            "text",
            "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>text</t></is></c>",
        };
    }
}
