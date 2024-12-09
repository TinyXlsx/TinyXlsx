namespace TinyXlsx;

public class Worksheet
{
    private readonly Stream stream;
    private readonly Workbook workbook;
    private int? rowIndex;

    public int Id { get; }

    public string Name { get; }

    public string RelationshipId { get; }

    public Worksheet(
        Workbook workbook,
        Stream stream,
        int id,
        string name,
        string relationshipId)
    {
        this.stream = stream;
        this.workbook = workbook;
        Id = id;
        Name = name;
        RelationshipId = relationshipId;
    }

    internal void BeginSheet()
    {
        Buffer.Append(stream, """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                mc:Ignorable="x14ac xr xr2 xr3"
                xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
                xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
                xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
                <dimension ref="A1:B1"/>
                <sheetViews>
                    <sheetView tabSelected="1" showRuler="1" showOutlineSymbols="1" defaultGridColor="1" colorId="64" zoomScale="100" workbookViewId="0"></sheetView>
                </sheetViews>
                <sheetFormatPr defaultColWidth="8.43" defaultRowHeight="15"/>
                <sheetData>
            """);
    }

    internal void EndSheet()
    {
        Buffer.Append(stream, """
                </sheetData>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>
            """);
        Buffer.Commit(stream);
        stream.Flush();
        stream.Close();
    }

    public void BeginRow(int rowIndex)
    {
        this.rowIndex = rowIndex + 1;

        Buffer.Append(stream, "<row r=\"");
        Buffer.Append(stream, this.rowIndex.Value);
        Buffer.Append(stream, "\">");
    }

    private void VerifyRowIsOpen()
    {
        if (rowIndex == null)
        {
            throw new InvalidOperationException($"A cell value can only be written after creating a row with {nameof(BeginRow)}.");
        }
    }

    public void EndRow()
    {
        Buffer.Append(stream, "</row>");
        rowIndex = null;
    }

    public void WriteCellValue(
        int columnIndex,
        double value)
    {
        VerifyRowIsOpen();

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, (char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        Buffer.Append(stream, rowIndex.Value);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</v></c>");
    }

    public void WriteCellValue(
        int columnIndex,
        double value,
        string format)
    {
        VerifyRowIsOpen();

        var numberFormatIndex = workbook.GetOrCreateNumberFormat(format);
        numberFormatIndex -= 163; // TODO: perhaps not the cleanest way of doing this, necessary for now to match Excel's numbering.

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, (char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        Buffer.Append(stream, rowIndex.Value);
        Buffer.Append(stream, "\" s=\"");
        Buffer.Append(stream, numberFormatIndex);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</v></c>");
    }

    public void WriteCellValue(
       int columnIndex,
       string value)
    {
        if (string.IsNullOrEmpty(value)) return;

        VerifyRowIsOpen();

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, (char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        Buffer.Append(stream, rowIndex.Value);
        Buffer.Append(stream, "\" t=\"inlineStr\"><is><t>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</t></is></c>");
    }

    public void WriteCellValue(
       int columnIndex,
       DateTime value)
    {
        WriteCellValue(
            columnIndex,
            value,
            "yyyy-MM-dd");
    }

    public void WriteCellValue(
       int columnIndex,
       DateTime value,
       string format)
    {
        VerifyRowIsOpen();

        var daysSinceBaseDate = (value - Constants.MinimumDate).TotalDays;

        if (daysSinceBaseDate < 0)
        {
            throw new NotSupportedException("The xlsx format does not support dates before 1990-01-01. Please write the value as a string instead.");
        }

        var numberFormatIndex = workbook.GetOrCreateNumberFormat(format);
        numberFormatIndex -= 163; // TODO: perhaps not the cleanest way of doing this, necessary for now to match Excel's numbering.

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, (char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        Buffer.Append(stream, rowIndex.Value);
        Buffer.Append(stream, "\" s=\"");
        Buffer.Append(stream, numberFormatIndex);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, daysSinceBaseDate);
        Buffer.Append(stream, "</v></c>");
    }
}
