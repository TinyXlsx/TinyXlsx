using System.IO.Compression;

namespace TinyXlsx;

public class Workbook : IDisposable
{
    private readonly Stream stream;
    private readonly ZipArchive archive;
    private readonly IList<Worksheet> worksheets;
    private readonly IDictionary<string, int> numberFormats;
    private readonly CompressionLevel compressionLevel;
    private bool disposedValue;

    public Workbook(
        string filePath,
        CompressionLevel compressionLevel = CompressionLevel.Optimal)
    {
        worksheets = new List<Worksheet>();
        numberFormats = new Dictionary<string, int>();
        this.compressionLevel = compressionLevel;

        stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
    }

    public Workbook(
        int capacity = 1024 * 1024,
        CompressionLevel compressionLevel = CompressionLevel.Optimal)
    {
        worksheets = new List<Worksheet>();
        numberFormats = new Dictionary<string, int>();
        this.compressionLevel = compressionLevel;

        stream = new MemoryStream(capacity);
        archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
    }

    public Stream Close()
    {
        AddRels();
        AddContentTypesXml();
        AddDocPropsAppXml();
        AddDocPropsCoreXml();
        AddWorkbookXml();
        AddStylesXml();
        AddSharedStringsXml();
        AddWorkbookXmlRels();

        archive.Dispose();
        stream.Position = 0;

        return stream;
    }

    public int GetOrCreateNumberFormat(string format)
    {
        if (numberFormats.TryGetValue(format, out var index))
        {
            return index;
        }

        index = numberFormats.Count + 164;
        numberFormats.Add(format, index);
        return index;
    }

    private void AddRels()
    {
        var entry = archive.CreateEntry("_rels/.rels", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
            <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
        </Relationships>
        """);
        Buffer.Commit(entryStream);
    }

    private void AddDocPropsAppXml()
    {
        var entry = archive.CreateEntry("docProps/app.xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="utf-8"?>
        <Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <ScaleCrop>false</ScaleCrop>
            <LinksUpToDate>false</LinksUpToDate>
            <SharedDoc>false</SharedDoc>
            <HyperlinksChanged>false</HyperlinksChanged>
            <Application>TinyXlsx</Application>
            <DocSecurity>0</DocSecurity>
        </Properties>
        """);
        Buffer.Commit(entryStream);
    }

    private void AddDocPropsCoreXml()
    {
        var entry = archive.CreateEntry("docProps/core.xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="utf-8"?>
        <coreProperties
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dc="http://purl.org/dc/elements/1.1/"
            xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
            xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
            <dcterms:created xsi:type="dcterms:W3CDTF">
        """);
        Buffer.Append(entryStream, DateTime.UtcNow.ToString("yyyy-MM-ddThh:mm:ssZ"));
        Buffer.Append(entryStream, "</dcterms:created><dc:creator></dc:creator></coreProperties>");
        Buffer.Commit(entryStream);
    }

    private void AddContentTypesXml()
    {
        var entry = archive.CreateEntry("[Content_Types].xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="utf-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
            <Default Extension="xml" ContentType="application/xml" />
            <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
            <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
            <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
            <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
        </Types>
        """);
        Buffer.Commit(entryStream);
    }

    private void AddStylesXml()
    {
        var entry = archive.CreateEntry("xl/styles.xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac x16r2 xr"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
            <numFmts count="
        """);

        Buffer.Append(entryStream, numberFormats.Count);
        Buffer.Append(entryStream, "\">");
        foreach (var item in numberFormats)
        {
            Buffer.Append(entryStream, "<numFmt numFmtId=\"");
            Buffer.Append(entryStream, item.Value);
            Buffer.Append(entryStream, "\" formatCode=\"");
            Buffer.Append(entryStream, item.Key);
            Buffer.Append(entryStream, "\"/>");
        }

        Buffer.Append(entryStream, """
            </numFmts>
            <fonts count="1">
                <font><sz val="11"/><color indexed="8"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
            </fonts>
            <fills count="2">
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="darkGray"/></fill>
            </fills>
            <borders count="1">
                <border><left/><right/><top/><bottom/><diagonal/></border>
            </borders>
            <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
            <cellXfs count="
        """);

        Buffer.Append(entryStream, numberFormats.Count + 1);
        Buffer.Append(entryStream, "\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        foreach (var item in numberFormats)
        {
            Buffer.Append(entryStream, "<xf numFmtId=\"");
            Buffer.Append(entryStream, item.Value);
            Buffer.Append(entryStream, "\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
        }

        Buffer.Append(entryStream, """
            </cellXfs>
        </styleSheet>
        """);
        Buffer.Commit(entryStream);
    }

    private void AddWorkbookXmlRels()
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
        </Relationships>
        """);
        Buffer.Commit(entryStream);
    }

    private void AddSharedStringsXml()
    {
        var entry = archive.CreateEntry("xl/sharedStrings.xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
        <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
            <si><t xml:space="preserve"></t></si>
            <si><t xml:space="preserve">Number</t></si>
        </sst>
        """);
        Buffer.Commit(entryStream);
    }

    public Worksheet BeginSheet(
        int id,
        string name,
        string relationshipId)
    {
        var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", compressionLevel);
        var entryStream = entry.Open();

        var worksheet = new Worksheet(
            this,
            entryStream,
            id,
            name,
            relationshipId);
        worksheet.BeginSheet();
        worksheets.Add(worksheet);
        return worksheet;
    }

    public Worksheet BeginSheet()
    {
        var name = $"Sheet{worksheets.Count + 1}";
        var id = worksheets.Count + 1;
        var relationshipId = $"rId{worksheets.Count + 3}";

        return BeginSheet(
            id,
            name,
            relationshipId);
    }

    public void EndSheet()
    {
        worksheets[worksheets.Count - 1].EndSheet();
    }

    private void AddWorkbookXml()
    {
        var entry = archive.CreateEntry("xl/workbook.xml", compressionLevel);
        using var entryStream = entry.Open();

        Buffer.Append(entryStream, """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                mc:Ignorable="x15 xr xr6 xr10 xr2"
                xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
                xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
                xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
                xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
                <workbookPr autoCompressPictures="1"/>
                <bookViews>
                    <workbookView tabRatio="600"/>
                </bookViews>
                <sheets>
            """);

        foreach (var worksheet in worksheets)
        {
            Buffer.Append(entryStream, "<sheet name=\"");
            Buffer.Append(entryStream, worksheet.Name);
            Buffer.Append(entryStream, "\" sheetId=\"");
            Buffer.Append(entryStream, worksheet.Id);
            Buffer.Append(entryStream, "\" r:id=\"");
            Buffer.Append(entryStream, worksheet.RelationshipId);
            Buffer.Append(entryStream, "\"></sheet>");
        }

        Buffer.Append(entryStream, """
                </sheets>
            </workbook>
            """);
        Buffer.Commit(entryStream);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposedValue) return;

        if (disposing)
        {
            archive.Dispose();
        }

        disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
