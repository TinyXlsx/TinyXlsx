using TinyXlsx;

namespace Tests;

[TestClass]
public class XlsxBuilderTests
{
    [DataRow("'", "&apos;")]
    [DataRow("\"", "&quot;")]
    [DataRow("&", "&amp;")]
    [DataRow("<", "&lt;")]
    [DataRow(">", "&gt;")]
    [DataRow("text", "text")]
    [DataRow("'\"&<>text'\"&<>", "&apos;&quot;&amp;&lt;&gt;text&apos;&quot;&amp;&lt;&gt;")]
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
