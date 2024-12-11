using System.Globalization;

namespace TinyXlsx;

public static class Constants
{
    public static readonly DateTime XlsxEpoch;
    public static readonly int MaximumCharactersPerCell;
    public static readonly int MaximumColumns;
    public static readonly DateTime MinimumDate;
    public static readonly int MaximumDoubleLength;
    public static readonly int MaximumIntegerLength;
    public static readonly int MaximumRows;
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

        // The XLSX format supports 65,490 styles, but that number includes built-in styles: 65,000 should be a safe margin.
        MaximumStyles = 65_000;
    }
}
