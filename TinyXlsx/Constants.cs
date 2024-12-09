using System.Globalization;

namespace TinyXlsx;

internal static class Constants
{
    internal static DateTime MinimumDate = new(1899, 12, 30);

    internal static int MaximumIntegerLength = int.MinValue.ToString(CultureInfo.InvariantCulture).Length;

    internal static int MaximumDoubleLength = double.MinValue.ToString(CultureInfo.InvariantCulture).Length;
}
