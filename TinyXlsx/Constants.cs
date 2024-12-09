using System.Globalization;

namespace TinyXlsx;

internal static class Constants
{
    internal static readonly int MaximumColumns;
    internal static readonly DateTime MinimumDate;
    internal static readonly int MaximumDoubleLength;
    internal static readonly int MaximumIntegerLength;
    internal static readonly int MaximumRows;

    static Constants()
    {
        MaximumColumns = 1 << 14;
        MinimumDate = new DateTime(1899, 12, 30);
        MaximumDoubleLength = double.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumIntegerLength = int.MinValue.ToString(CultureInfo.InvariantCulture).Length;
        MaximumRows = 1 << 20;
    }
}
