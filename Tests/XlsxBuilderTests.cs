using TinyXlsx;

namespace Tests;

[TestClass]
public class XlsxBuilderTests
{
    [DataRow("&", "&amp;")]
    [DataRow("<", "&lt;")]
    [TestMethod]
    public void SpecialCharacterIsCorrectlyEncoded(
        string character,
        string expected)
    {
        var memoryStream = new MemoryStream();
        var xlsxBuilder = new XlsxBuilder();

        xlsxBuilder.Append(memoryStream, character);
        xlsxBuilder.Commit(memoryStream);

        memoryStream.Position = 0;
        using var reader = new StreamReader(memoryStream);
        var actual = reader.ReadToEnd();
        Assert.AreEqual(expected, actual);
    }
}
