using BenchmarkDotNet.Attributes;
using LargeXlsx;

namespace Benchmarks;

public partial class Benchmarks
{
    [Benchmark]
    public void LargeXlsx()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream, SharpCompress.Compressors.Deflate.CompressionLevel.Default);
        var worksheet = xlsxWriter.BeginWorksheet("Sheet1");

        var numberStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimal);
        var percentageStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimalPercentage);
        var scientificStyle = XlsxStyle.Default.With(XlsxNumberFormat.Scientific);
        var currencyStyle1 = XlsxStyle.Default.With(new XlsxNumberFormat("$#,##0.00"));
        var currencyStyle2 = XlsxStyle.Default.With(new XlsxNumberFormat("#,##0.00 [$USD]")); 

        for (var i = 0; i < Records; i++)
        {
            worksheet
                .BeginRow()
                .Write(123.456)      
                .Write(DateTime.Now)
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
