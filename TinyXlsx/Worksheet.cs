namespace TinyXlsx;

/// <summary>
/// Represents a worksheet within a <see cref="Workbook"/> for writing data to an XLSX file. 
/// </summary>
public class Worksheet
{
    private readonly XlsxBuilder xlsxBuilder;
    private readonly Stream stream;
    private readonly Workbook workbook;
    private int lastWrittenRowIndex;
    private int lastWrittenColumnIndex;

    internal int Id { get; }
    internal string Name { get; }
    internal string RelationshipId { get; }

    public Worksheet(
        Workbook workbook,
        XlsxBuilder xlsxBuilder,
        Stream stream,
        int id,
        string name,
        string relationshipId)
    {
        this.workbook = workbook;
        this.xlsxBuilder = xlsxBuilder;
        this.stream = stream;
        Id = id;
        Name = name;
        RelationshipId = relationshipId;
    }

    /// <summary>
    /// Begins a new row within the worksheet at the specified index, automatically ending any previous row.
    /// </summary>
    /// <param name="rowIndex">
    /// The one-based index to start the row at.
    /// </param>
    public void BeginRowAt(int rowIndex)
    {
        EndRow();
        VerifyCanBeginRow(rowIndex);

        xlsxBuilder.Append(stream, "<row r=\"");
        xlsxBuilder.Append(stream, rowIndex);
        xlsxBuilder.Append(stream, "\">");

        lastWrittenRowIndex = rowIndex;
    }

    /// <summary>
    /// Begins a new row within the worksheet, automatically ending any previous row.
    /// </summary>
    public void BeginRow()
    {
        BeginRowAt(lastWrittenRowIndex + 1);
    }

    /// <summary>
    /// Writes a formula to the next cell.
    /// </summary>
    /// <param name="formula">
    /// The formula as a string value to write to the cell.
    /// </param>
    public void WriteCellFormula(string formula)
    {
        WriteCellFormulaAt(lastWrittenColumnIndex + 1, formula);
    }

    /// <summary>
    /// Writes a <see cref="bool"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="bool"/> value to write to the cell.
    /// </param>
    public void WriteCellFormulaAt(
        int columnIndex,
        string formula)
    {
        VerifyCanWriteCellValue(columnIndex);

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"e\"><f>");
        xlsxBuilder.Append(stream, formula);
        xlsxBuilder.Append(stream, "</f></c>");

        lastWrittenColumnIndex = columnIndex;
    }

    /// <summary>
    /// Writes a <see cref="bool"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="bool"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(bool value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="bool"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="bool"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
        int columnIndex,
        bool value)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"b\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="decimal"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="decimal"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(decimal value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="decimal"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="decimal"/> value to write to the cell.
    /// </param>
    /// <param name="format"></param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValue(
        decimal value,
        string format)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="decimal"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="decimal"/> value to write to the cell.
    /// </param>
    /// <param name="format">
    /// The number format to apply to the cell.
    /// </param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValueAt(
        int columnIndex,
        decimal value,
        string format)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        var (zeroBasedIndex, _) = workbook.GetOrCreateNumberFormat(format);

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" s=\"");
        xlsxBuilder.Append(stream, zeroBasedIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="decimal"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="decimal"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
        int columnIndex,
        decimal value)
    {
        VerifyCanWriteCellValue(columnIndex);
        lastWrittenColumnIndex = columnIndex;

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="double"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(double value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
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
        WriteCellValueAt(lastWrittenColumnIndex + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
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

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="double"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
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

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" s=\"");
        xlsxBuilder.Append(stream, zeroBasedIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    /// <summary>
    /// Writes a <see cref="int"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="int"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(int value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="int"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="int"/> value to write to the cell.
    /// </param>
    /// <param name="format"></param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValue(
        int value,
        string format)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="int"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="int"/> value to write to the cell.
    /// </param>
    /// <param name="format">
    /// The number format to apply to the cell.
    /// </param>
    /// <remarks>
    /// The specified format must be valid. Invalid formats may result in a repair prompt from the XLSX viewer.
    /// </remarks>
    public void WriteCellValueAt(
        int columnIndex,
        int value,
        string format)
    {
        VerifyCanWriteCellValue(columnIndex);

        var (zeroBasedIndex, _) = workbook.GetOrCreateNumberFormat(format);

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" s=\"");
        xlsxBuilder.Append(stream, zeroBasedIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");

        lastWrittenColumnIndex = columnIndex;
    }

    /// <summary>
    /// Writes a <see cref="int"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
    /// </param>
    /// <param name="value">
    /// The <see cref="int"/> value to write to the cell.
    /// </param>
    public void WriteCellValueAt(
        int columnIndex,
        int value)
    {
        VerifyCanWriteCellValue(columnIndex);

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</v></c>");

        lastWrittenColumnIndex = columnIndex;
    }

    /// <summary>
    /// Writes a <see cref="string"/> value to the next cell.
    /// </summary>
    /// <param name="value">
    /// The <see cref="string"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(string value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
    }

    /// <summary>
    /// Writes a <see cref="string"/> value to the specified cell.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
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

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" t=\"inlineStr\"><is><t>");
        xlsxBuilder.Append(stream, value);
        xlsxBuilder.Append(stream, "</t></is></c>");
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the next cell with a default format of yyyy-MM-dd.
    /// </summary>
    /// <param name="value">
    /// The <see cref="DateTime"/> value to write to the cell.
    /// </param>
    public void WriteCellValue(DateTime value)
    {
        WriteCellValueAt(lastWrittenColumnIndex + 1, value);
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
        WriteCellValueAt(lastWrittenColumnIndex + 1, value, format);
    }

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the specified cell with a default format of yyyy-MM-dd.
    /// </summary>
    /// <param name="columnIndex">
    /// The one-based column index of the cell to write to.
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
    /// The one-based column index of the cell to write to.
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

        xlsxBuilder.Append(stream, "<c r=\"");
        xlsxBuilder.AppendColumnKey(stream, columnIndex);
        xlsxBuilder.Append(stream, lastWrittenRowIndex);
        xlsxBuilder.Append(stream, "\" s=\"");
        xlsxBuilder.Append(stream, zeroBasedIndex);
        xlsxBuilder.Append(stream, "\" t=\"n\"><v>");
        xlsxBuilder.Append(stream, daysSinceEpoch);
        xlsxBuilder.Append(stream, "</v></c>");
    }

    internal void BeginSheet()
    {
        // If data has been written to the stream, then BeginSheet has already been called.
        // The stream does not support seeking, the only value which can be checked is Position.
        if (stream.Position > 0) return;

        // Intentionally leaving <dimension /> empty as stream does not support seeking.
        xlsxBuilder.Append(stream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<sheetViews>"
            + "<sheetView workbookViewId=\"0\"></sheetView>"
            + "</sheetViews>"
            + "<sheetData>");
    }

    internal void EndSheet()
    {
        // If the stream is closed, then EndSheet has already been called.
        if (!stream.CanWrite) return;

        EndRow();
        xlsxBuilder.Append(stream,
            "</sheetData>"
            + "</worksheet>");
        xlsxBuilder.Commit(stream);
        stream.Flush();
        stream.Close();
    }

    private void EndRow()
    {
        // It's possible that there is no row to end, if BeginRow hasn't been called yet.
        if (lastWrittenRowIndex == 0) return;

        VerifyCanEndRow();

        xlsxBuilder.Append(stream, "</row>");
        lastWrittenColumnIndex = 0;
    }

    private void VerifyCanBeginRow(int rowIndex)
    {
        if (rowIndex <= lastWrittenRowIndex)
        {
            throw new InvalidOperationException($"A row with an index equal to or higher than {rowIndex} was already written to.");
        }

        if (rowIndex > Constants.MaximumRows)
        {
            throw new InvalidOperationException($"The XLSX format only supports {Constants.MaximumRows} rows.");
        }

        if (rowIndex < 1)
        {
            throw new InvalidOperationException("The XLSX format supports a minmium row index of 1.");
        }
    }

    private void VerifyCanEndRow()
    {
        // No need to guard against empty rows, i.e. <row></row>, as the XLSX format allows it.

        if (lastWrittenRowIndex == 0)
        {
            throw new InvalidOperationException($"A row cannot be closed before it was opened with {nameof(BeginRow)}.");
        }
    }

    private void VerifyCanWriteCellValue(int columnIndex)
    {
        if (lastWrittenRowIndex == 0)
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

        if (columnIndex < 1)
        {
            throw new InvalidOperationException("The XLSX format supports a minmium column index of 1.");
        }
    }
}
