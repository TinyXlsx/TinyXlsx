namespace TinyXlsx;

public class Worksheet : IDisposable
{
    private readonly Workbook workbook;
    private readonly Stream stream;
    private int? rowIndex;
    private bool disposedValue;

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
        this.workbook = workbook;
        this.stream = stream;
        Id = id;
        Name = name;
        RelationshipId = relationshipId;
    }

    internal async Task BeginSheetAsync()
    {
        stream.BufferPooledWrite("""
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

    internal async Task EndSheetAsync()
    {
        stream.BufferPooledWrite("""
                </sheetData>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>
            """);
        await stream.FlushAsync();
        stream.Close();
    }

    public async Task BeginRowAsync(int rowIndex)
    {
        this.rowIndex = rowIndex + 1;
        stream.BufferPooledWrite("<row r=\"");
        stream.BufferPooledWrite(this.rowIndex.Value);
        stream.BufferPooledWrite("\">");
    }

    private void VerifyRowIsOpen()
    {
        if (rowIndex == null)
        {
            throw new InvalidOperationException($"A cell value can only be written after creating a row with {nameof(BeginRowAsync)}.");
        }
    }

    public async Task EndRowAsync()
    {
        stream.BufferPooledWrite("</row>");
        rowIndex = null;
    }

    public void WriteCellValue(
        int columnIndex,
        double value)
    {
        VerifyRowIsOpen();

        stream.BufferPooledWrite("<c r=\"");
        stream.BufferPooledWrite((char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        stream.BufferPooledWrite(rowIndex.Value);
        stream.BufferPooledWrite("\" t=\"n\"><v>");
        stream.BufferPooledWrite(value);
        stream.BufferPooledWrite("</v></c>");
    }

    public void WriteCellValue(
        int columnIndex,
        double value,
        string format)
    {
        VerifyRowIsOpen();

        var numberFormatIndex = workbook.GetOrCreateNumberFormat(format);
        numberFormatIndex -= 163; // TODO: perhaps not the cleanest way of doing this, necessary for now to match Excel's numbering.

        stream.BufferPooledWrite("<c r=\"");
        stream.BufferPooledWrite((char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        stream.BufferPooledWrite(rowIndex.Value);
        stream.BufferPooledWrite("\" s=\"");
        stream.BufferPooledWrite(numberFormatIndex);
        stream.BufferPooledWrite("\" t=\"n\"><v>");
        stream.BufferPooledWrite(value);
        stream.BufferPooledWrite("</v></c>");
    }

    public void WriteCellValue(
       int columnIndex,
       string value)
    {
        VerifyRowIsOpen();

        stream.BufferPooledWrite("<c r=\"");
        stream.BufferPooledWrite((char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        stream.BufferPooledWrite(rowIndex.Value);
        stream.BufferPooledWrite("\" t=\"inlineStr\"><is><t>");
        stream.BufferPooledWrite(value);
        stream.BufferPooledWrite("</t></is></c>");
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

        stream.BufferPooledWrite("<c r=\"");
        stream.BufferPooledWrite((char)('A' + columnIndex)); // TODO: handle more than 26 columns.
        stream.BufferPooledWrite(rowIndex.Value);
        stream.BufferPooledWrite("\" s=\"");
        stream.BufferPooledWrite(numberFormatIndex);
        stream.BufferPooledWrite("\" t=\"n\"><v>");
        stream.BufferPooledWrite(daysSinceBaseDate);
        stream.BufferPooledWrite("</v></c>");
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposedValue) return;

        if (disposing)
        {
        }

        disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
