using TinyXlsx;

namespace Tests;

[TestClass]
public class ColumnKeyCacheTests
{
    [DataRow(1, "A")]
    [DataRow(27, "AA")]
    [DataRow(280, "JT")]
    [DataRow(703, "AAA")]
    [DataRow(26, "Z")]
    [DataRow(702, "ZZ")]
    [DataRow(16_384, "XFD")]
    [TestMethod]
    public void GetKeyReturnsCorrectKey(int columnIndex, string  expected)
    {
        var actual = ColumnKeyCache.GetKey(columnIndex);

        Assert.AreEqual(expected, actual);
    }
}
