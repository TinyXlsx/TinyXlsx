using System.Globalization;

namespace TinyXlsx;

internal static class Constants
{
    internal static readonly int MaximumCharactersPerCell;
    internal static readonly int MaximumColumns;
    internal static readonly DateTime MinimumDate;
    internal static readonly int MaximumDoubleLength;
    internal static readonly int MaximumIntegerLength;
    internal static readonly int MaximumRows;
    internal static readonly int MaximumStyles;

    static Constants()
    {
        MaximumCharactersPerCell = 32_767;
        MaximumColumns = 16_384;
        MinimumDate = new DateTime(1899, 12, 30);
        MaximumDoubleLength = double.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumIntegerLength = int.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumRows = 1_048_576;

        // The XLSX format supports 65,490 styles, but that number includes built-in styles: 65,000 should be a safe margin.
        MaximumStyles = 65_000;
    }
}
