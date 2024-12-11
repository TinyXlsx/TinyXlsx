using System.Globalization;

namespace TinyXlsx;

/// <summary>
/// Represents a set of constants for the XLSX format.
/// </summary>
public static class Constants
{
    /// <summary>
    /// The epoch date used in the XLSX format. Any dates are saved as an offset based on this date.
    /// </summary>
    public static readonly DateTime XlsxEpoch;

    /// <summary>
    /// The maximum number of characters allowed in a single cell.
    /// </summary>
    public static readonly int MaximumCharactersPerCell;

    /// <summary>
    /// The maximum number of columns allowed in an XLSX worksheet.
    /// </summary>
    public static readonly int MaximumColumns;

    /// <summary>
    /// The earliest date that is represented correctly in an XLSX viewer.
    /// </summary>
    public static readonly DateTime MinimumDate;

    /// <summary>
    /// The maximum number of characters required to write a <see cref="double"/> value as a string.
    /// </summary>
    public static readonly int MaximumDoubleLength;

    /// <summary>
    /// The maximum number of characters required to write a <see cref="int"/> value as a string.
    /// </summary>
    public static readonly int MaximumIntegerLength;

    /// <summary>
    /// The maximum number of rows allowed in an XLSX worksheet.
    /// </summary>
    public static readonly int MaximumRows;

    /// <summary>
    /// The maximum number of styles supported in an XLSX workbook.
    /// </summary>
    /// <remarks>
    /// The XLSX format supports 65,490 styles, but that number includes built-in styles: 65,000 is used as a safe margin.
    /// </remarks>
    public static readonly int MaximumStyles;

    static Constants()
    {
        MaximumCharactersPerCell = 32_767;
        MaximumColumns = 16_384;
        XlsxEpoch = new DateTime(1899, 12, 30);
        MinimumDate = new DateTime(1900, 1, 1);
        MaximumDoubleLength = double.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumIntegerLength = int.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumRows = 1_048_576;
        MaximumStyles = 65_000;
    }
}
