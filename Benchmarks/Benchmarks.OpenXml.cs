using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace Benchmarks;

public partial class Benchmarks
{
    [Benchmark]
    public void OpenXml()
    {
        using var memoryStream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()),
            new Fills(new Fill()),
            new Borders(new Border()),
            new CellFormats(
                new CellFormat(),
                new CellFormat
                {
                    NumberFormatId = 164,
                    ApplyNumberFormat = true,
                },
                new CellFormat
                {
                    NumberFormatId = 165,
                    ApplyNumberFormat = true,
                },
                new CellFormat
                {
                    NumberFormatId = 166,
                    ApplyNumberFormat = true,
                },
                new CellFormat
                {
                    NumberFormatId = 167,
                    ApplyNumberFormat = true,
                },
                new CellFormat
                {
                    NumberFormatId = 168,
                    ApplyNumberFormat = true,
                }
            ),
            new NumberingFormats(
                new NumberingFormat
                {
                    NumberFormatId = 164,
                    FormatCode = "0.00",
                },
                new NumberingFormat
                {
                    NumberFormatId = 165,
                    FormatCode = "0.00%",
                },
                new NumberingFormat
                {
                    NumberFormatId = 166,
                    FormatCode = "0.00E+00",
                },
                new NumberingFormat
                {
                    NumberFormatId = 167,
                    FormatCode = "$#,##0.00",
                },
                new NumberingFormat
                {
                    NumberFormatId = 168,
                    FormatCode = "#,##0.00 [$USD]",
                }
            )
        );
        stylesPart.Stylesheet.Save();

        for (var i = 0; i < Records; i++)
        {
            var dataRow = new Row();

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Date,
                CellValue = new CellValue(DateTime.Now),
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.InlineString,
                CellValue = new CellValue("Text"),
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
                StyleIndex = 1,
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
                StyleIndex = 2,
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
                StyleIndex = 3,
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
                StyleIndex = 4,
            });

            dataRow.Append(new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(123.456),
                StyleIndex = 5,
            });

            sheetData.Append(dataRow);
        }

        var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
        };
        sheets.Append(sheet);

        workbookPart.Workbook.Save();
    }
}
