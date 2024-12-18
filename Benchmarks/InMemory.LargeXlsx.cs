using BenchmarkDotNet.Attributes;
using LargeXlsx;

namespace Benchmarks;

public partial class InMemory
{
    [Benchmark]
    public void LargeXlsx()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        var worksheet = xlsxWriter.BeginWorksheet("Sheet1");

        var dateTimeUtcStyle = XlsxStyle.Default.With(new XlsxNumberFormat("yyyy-MM-dd"));
        var dateTimeAlternativeStyle = XlsxStyle.Default.With(new XlsxNumberFormat("yyyy/MM/dd"));
        var numberStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimal);
        var percentageStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimalPercentage);
        var scientificStyle = XlsxStyle.Default.With(XlsxNumberFormat.Scientific);
        var currencyStyle1 = XlsxStyle.Default.With(new XlsxNumberFormat("$#,##0.00"));
        var currencyStyle2 = XlsxStyle.Default.With(new XlsxNumberFormat("#,##0.00 [$USD]"));

        for (var i = 0; i < Records; i++)
        {
            worksheet
                .BeginRow()
                .Write(false)
                .Write(123456)
                .Write(123.456m)
                .Write(123.456)
                .Write(DateTime.Now, dateTimeUtcStyle)
                .Write(DateTime.Now, dateTimeAlternativeStyle)
                .Write("Text")
                .Write(123.456, numberStyle)
                .Write(123.456, percentageStyle)
                .Write(123.456, scientificStyle)
                .Write(123.456, currencyStyle1)
                .Write(123.456, currencyStyle2);
        }
        xlsxWriter.Dispose();
    }
}
