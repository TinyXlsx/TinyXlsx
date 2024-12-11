namespace TinyXlsx;

/// <summary>
/// Represents a worksheet within a <see cref="Workbook"/> for writing data to an XLSX file. 
/// </summary>
public class Worksheet
{
    private readonly Stream stream;
    private readonly Workbook workbook;
    private int? lastWrittenRowIndex;
    private int? lastWrittenColumnIndex;
    private int? internalRowIndex;

    internal int Id { get; }
    internal string Name { get; }
    internal string RelationshipId { get; }

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

    /// <summary>
    /// Begins a new row within the worksheet at the specified index, automatically ending any previous row.
    /// </summary>
    /// <param name="rowIndex">
    /// The zero-based index to start the row at.
    /// </param>
    public void BeginRowAt(int rowIndex)
    {
        EndRow();
        VerifyCanBeginRow(rowIndex);

        internalRowIndex = rowIndex + 1;
        lastWrittenRowIndex = rowIndex;

        Buffer.Append(stream, "<row r=\"");
        Buffer.Append(stream, internalRowIndex.Value);
        Buffer.Append(stream, "\">");
    }

    /// <summary>
    /// Begins a new row within the worksheet, automatically ending any previous row.
    /// </summary>
    public void BeginRow()
    {
        BeginRowAt((lastWrittenRowIndex ?? -1) + 1);
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="double"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(double value)
    {
        WriteCellValueAt((lastWrittenColumnIndex ?? -1) + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="double"/> value to write to the cell.
    /// </param>
    /// <param name="format"></param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValue(
        double value,
        string format)
    {
        WriteCellValueAt((lastWrittenColumnIndex ?? -1) + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The zero-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="double"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
        int columnIndex,
        double value)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, ColumnKeyCache.GetKey(columnIndex));
        Buffer.Append(stream, internalRowIndex!.Value);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The zero-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="double"/> value to write to the cell.
    /// </param>
    /// <param name="format">
    /// The number format to apply to the cell.
    /// </param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValueAt(
        int columnIndex,
        double value,
        string format)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        var (zeroBasedIndex, _) = workbook.GetOrCreateNumberFormat(format);

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, ColumnKeyCache.GetKey(columnIndex));
        Buffer.Append(stream, internalRowIndex!.Value);
        Buffer.Append(stream, "\" s=\"");
        Buffer.Append(stream, zeroBasedIndex);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="string"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="string"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(string value)
    {
        WriteCellValueAt((lastWrittenColumnIndex ?? -1) + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="string"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The zero-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="string"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
       int columnIndex,
       string value)
    {
        if (string.IsNullOrEmpty(value)) return;
        if (value.Length > Constants.MaximumCharactersPerCell)
        {
            throw new NotSupportedException($"The XLSX format does not support more than {Constants.MaximumCharactersPerCell} characters in a cell.");
        }

        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, ColumnKeyCache.GetKey(columnIndex));
        Buffer.Append(stream, internalRowIndex!.Value);
        Buffer.Append(stream, "\" t=\"inlineStr\"><is><t>");
        Buffer.Append(stream, value);
        Buffer.Append(stream, "</t></is></c>");
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the next cell with a default format of yyyy-MM-dd.
    /// </summary>
    /// <param name="value">
    /// The <see cref="DateTime"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(DateTime value)
    {
        WriteCellValueAt((lastWrittenColumnIndex ?? -1) + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="DateTime"/> value to write to the cell.
    /// </param>
    /// <param name="format">
    /// The date format to apply to the cell.
    /// </param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValue(
        DateTime value,
        string format)
    {
        WriteCellValueAt((lastWrittenColumnIndex ?? -1) + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the specified cell with a default format of yyyy-MM-dd.
    /// </summary>
    /// <param name="columnIndex">
    /// The zero-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="DateTime"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
       int columnIndex,
       DateTime value)
    {
        WriteCellValueAt(
            columnIndex,
            value,
            "yyyy-MM-dd");
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The zero-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="DateTime"/> value to write to the cell.
    /// </param>
    /// <param name="format">
    /// The date format to apply to the cell.
    /// </param>
    /// <exception cref="NotSupportedException">
    /// Thrown if the specified date is before 1990-01-01, which is unsupported by the XLSX format.
    /// </exception>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValueAt(
       int columnIndex,
       DateTime value,
       string format)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        if (value < Constants.MinimumDate)
        {
            throw new NotSupportedException("The XLSX format does not support dates before 1990-01-01. Consider writing the value as a number or string instead.");
        }

        // Account for leap year bug.
        if (value < Constants.LeapYearBugCorrectionDate)
        {
            value = value.AddDays(-1);
        }

        var daysSinceEpoch = (value - Constants.XlsxEpoch).TotalDays;

        var (zeroBasedIndex, _) = workbook.GetOrCreateNumberFormat(format);

        Buffer.Append(stream, "<c r=\"");
        Buffer.Append(stream, ColumnKeyCache.GetKey(columnIndex));
        Buffer.Append(stream, internalRowIndex!.Value);
        Buffer.Append(stream, "\" s=\"");
        Buffer.Append(stream, zeroBasedIndex);
        Buffer.Append(stream, "\" t=\"n\"><v>");
        Buffer.Append(stream, daysSinceEpoch);
        Buffer.Append(stream, "</v></c>");
    }

    internal void BeginSheet()
    {
        // If data has been written to the stream, then BeginSheet has already been called.
        // The stream does not support seeking, the only value which can be checked is Position.
        if (stream.Position > 0) return;

        // Intentionally leaving <dimension /> empty as stream does not support seeking.
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
                <sheetViews>
                    <sheetView tabSelected="1" showRuler="1" showOutlineSymbols="1" defaultGridColor="1" colorId="64" zoomScale="100" workbookViewId="0"></sheetView>
                </sheetViews>
                <sheetFormatPr defaultColWidth="8.43" defaultRowHeight="15"/>
                <sheetData>
            """);
    }

    internal void EndSheet()
    {
        // If the stream is closed, then EndSheet has already been called.
        if (!stream.CanWrite) return;

        EndRow();
        Buffer.Append(stream, """
                </sheetData>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>
            """);
        Buffer.Commit(stream);
        stream.Flush();
        stream.Close();
    }

    private void EndRow()
    {
        // It's possible that there is no row to end, if BeginRow hasn't been called yet.
        if (lastWrittenRowIndex == null) return;

        VerifyCanEndRow();

        Buffer.Append(stream, "</row>");
        internalRowIndex = null;
        lastWrittenColumnIndex = null;
    }

    private void VerifyCanBeginRow(int rowIndex)
    {
        if (internalRowIndex != null)
        {
            throw new InvalidOperationException($"A new row cannot be started until the previous row was closed with {nameof(EndRow)}.");
        }

        if (rowIndex <= lastWrittenRowIndex)
        {
            throw new InvalidOperationException($"A row with an index equal to or higher than {rowIndex} was already written to.");
        }

        if (rowIndex > Constants.MaximumRows)
        {
            throw new InvalidOperationException($"The XLSX format only supports {Constants.MaximumRows} rows.");
        }

        if (rowIndex < 0)
        {
            throw new InvalidOperationException($"The XLSX format does not support a negative row index.");
        }
    }

    private void VerifyCanEndRow()
    {
        // No need to guard against empty rows, i.e. <row></row>, as the XLSX format allows it.

        if (internalRowIndex == null)
        {
            throw new InvalidOperationException($"A row cannot be closed before it was opened with {nameof(BeginRow)}.");
        }
    }

    private void VerifyCanWriteCellValue(int columnIndex)
    {
        if (internalRowIndex == null)
        {
            throw new InvalidOperationException($"A cell value can only be written after creating a row with {nameof(BeginRow)}.");
        }

        if (columnIndex <= lastWrittenColumnIndex)
        {
            throw new InvalidOperationException($"A cell with an index equal to or higher than {columnIndex} was already written to.");
        }

        if (columnIndex > Constants.MaximumColumns)
        {
            throw new InvalidOperationException($"The XLSX format only supports {Constants.MaximumColumns} columns.");
        }

        if (columnIndex < 0)
        {
            throw new InvalidOperationException($"The XLSX format does not support a negative column index.");
        }
    }
}
