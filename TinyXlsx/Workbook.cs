using System.IO.Compression;

namespace TinyXlsx;

/// <summary>
/// Represents an in-memory workbook for creating an XLSX file.
/// </summary>
public class Workbook : IDisposable
{
    private readonly XlsxBuilder xlsxBuilder;
    private readonly Stream stream;
    private readonly ZipArchive archive;
    private readonly List<Worksheet> worksheets;
    private readonly Stylesheet stylesheet;
    private readonly CompressionLevel compressionLevel;
    private bool disposedValue;

    /// <summary>
    /// Initializes a new instance of the <see cref="Workbook"/> class writing to a file.
    /// </summary>
    /// <param name="filePath">
    /// The relative or absolute path of the file. The XLSX format does not support file paths exceeding 218 characters.
    /// </param>
    /// <param name="compressionLevel">
    /// The level of compression to apply to the workbook.
    /// Setting this lower will consume less resources but result in larger files.
    /// Setting this higher will consume more resources but result in smaller files.
    /// </param>
    public Workbook(
        string filePath,
        CompressionLevel compressionLevel = CompressionLevel.Fastest)
    {
        // No need to guard against filePath exceeding maximum length, as the XLSX viewer throws an error when opening the file.

        worksheets = [];
        stylesheet = new Stylesheet();
        this.compressionLevel = compressionLevel;

        stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
        xlsxBuilder = new XlsxBuilder();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Workbook"/> class writing to a <see cref="MemoryStream"/>.
    /// </summary>
    /// <param name="capacity">
    /// The initial size of the internal array in bytes. Consider setting this to a higher value than the resulting file size.
    /// </param>
    /// <param name="compressionLevel">
    /// The level of compression to apply to the workbook.
    /// Setting this lower will consume less resources but result in larger files.
    /// Setting this higher will consume more resources but result in smaller files.
    /// </param>
    public Workbook(
        int capacity = 1024 * 64,
        CompressionLevel compressionLevel = CompressionLevel.Fastest)
    {
        worksheets = [];
        stylesheet = new Stylesheet();
        this.compressionLevel = compressionLevel;

        stream = new MemoryStream(capacity);
        archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
        xlsxBuilder = new XlsxBuilder();
    }

    /// <summary>
    /// Begins a new worksheet with the specified name within the workbook, automatically ending any previously active worksheet.
    /// </summary>
    /// <param name="name">
    /// The name of the worksheet.
    /// </param>
    /// <returns>
    /// A <see cref="Worksheet"/> instance representing the new worksheet.
    /// </returns>
    public Worksheet BeginSheet(string name)
    {
        var id = worksheets.Count + 1;
        var relationshipId = $"rId{worksheets.Count + 3}";

        return BeginSheet(
            id,
            name,
            relationshipId);
    }

    /// <summary>
    /// Begins a new worksheet within the workbook with an automatically generated name, automatically ending any previously active worksheet.
    /// </summary>
    /// <returns>
    /// A <see cref="Worksheet"/> instance representing the new worksheet.
    /// </returns>
    public Worksheet BeginSheet()
    {
        var id = worksheets.Count + 1;
        var name = $"Sheet{id}";
        var relationshipId = $"rId{worksheets.Count + 3}";

        return BeginSheet(
            id,
            name,
            relationshipId);
    }

    /// <summary>
    /// Writes the final data to the <see cref="Stream"/>.
    /// If the <see cref="Workbook"/> is writing to a <see cref="FileStream"/>, the stream is disposed.
    /// If the <see cref="Workbook"/> is writing to a <see cref="MemoryStream"/>, its position is set to 0.
    /// </summary>
    /// <returns>
    /// The underlying <see cref="Stream"/> containing the workbook data.
    /// </returns>
    public Stream Close()
    {
        EndSheet();
        AddRels();
        AddContentTypesXml();
        AddDocPropsAppXml();
        AddDocPropsCoreXml();
        AddWorkbookXml();
        AddStylesXml();
        AddSharedStringsXml();
        AddWorkbookXmlRels();

        archive.Dispose();
        if (stream is FileStream)
        {
            stream.Dispose();
        }
        else
        {
            stream.Position = 0;
        }

        return stream;
    }

    /// <summary>
    /// Disposes the <see cref="Workbook"/> and releases its resources.
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes the resources.
    /// </summary>
    /// <param name="disposing">
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (disposedValue) return;

        if (disposing)
        {
            archive.Dispose();
            stream.Dispose();
        }

        disposedValue = true;
    }

    private void AddContentTypesXml()
    {
        var entry = archive.CreateEntry("[Content_Types].xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"u8
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"u8
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"u8
            + "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"u8
            + "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"u8
            + "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />"u8
            + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"u8
            + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"u8
            + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"u8);

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<Override PartName=\"/xl/worksheets/sheet"u8);
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />"u8);
        }
        xlsxBuilder.Append(entryStream, "</Types>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddDocPropsAppXml()
    {
        var entry = archive.CreateEntry("docProps/app.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8
            + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"u8
            + "<Application>TinyXlsx 0.1.0</Application>"u8
            + "<AppVersion>15.0000</AppVersion>"u8
            + "</Properties>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddDocPropsCoreXml()
    {
        var entry = archive.CreateEntry("docProps/core.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8
            + "<coreProperties "u8
            + "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""u8
            + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\""u8
            + "xmlns:dcterms=\"http://purl.org/dc/terms/\""u8
            + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\""u8
            + "xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">"u8
            + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">"u8);

        xlsxBuilder.Append(entryStream, DateTime.UtcNow.ToString("yyyy-MM-ddThh:mm:ssZ"));
        xlsxBuilder.Append(entryStream, "</dcterms:created><dc:creator></dc:creator></coreProperties>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddRels()
    {
        var entry = archive.CreateEntry("_rels/.rels", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"u8
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\" />"u8
            + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"u8
            + "</Relationships>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddSharedStringsXml()
    {
        var entry = archive.CreateEntry("xl/sharedStrings.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>"u8
            + "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8
            + "</sst>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddStylesXml()
    {
        var entry = archive.CreateEntry("xl/styles.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8
            + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8
            + "<numFmts count=\""u8);

        xlsxBuilder.Append(entryStream, stylesheet.Formats.Count);
        xlsxBuilder.Append(entryStream, "\">"u8);
        foreach (var item in stylesheet.Formats)
        {
            xlsxBuilder.Append(entryStream, "<numFmt numFmtId=\""u8);
            xlsxBuilder.Append(entryStream, item.Value.CustomFormatIndex);
            xlsxBuilder.Append(entryStream, "\" formatCode=\""u8);
            xlsxBuilder.Append(entryStream, item.Key);
            xlsxBuilder.Append(entryStream, "\"/>"u8);
        }

        xlsxBuilder.Append(entryStream,
            "</numFmts>"u8
            + "<fonts count=\"1\">"u8
            + "<font><sz val=\"11\"/><color indexed=\"8\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"u8
            + "</fonts>"u8
            + "<fills count=\"2\">"u8
            + "<fill><patternFill patternType=\"none\"/></fill>"u8
            + "<fill><patternFill patternType=\"darkGray\"/></fill>"u8
            + "</fills>"u8
            + "<borders count=\"1\">"u8
            + "<border><left/><right/><top/><bottom/><diagonal/></border>"u8
            + "</borders>"u8
            + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"u8
            + "<cellXfs count=\""u8);

        xlsxBuilder.Append(entryStream, stylesheet.Formats.Count + 1);
        xlsxBuilder.Append(entryStream, "\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"u8);
        foreach (var item in stylesheet.Formats)
        {
            xlsxBuilder.Append(entryStream, "<xf numFmtId=\""u8);
            xlsxBuilder.Append(entryStream, item.Value.CustomFormatIndex);
            xlsxBuilder.Append(entryStream, "\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"u8);
        }

        xlsxBuilder.Append(entryStream,
            "</cellXfs>"u8
            + "</styleSheet>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddWorkbookXml()
    {
        var entry = archive.CreateEntry("xl/workbook.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8
            + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8
            + "<sheets>"u8);

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<sheet name=\""u8);
            xlsxBuilder.Append(entryStream, worksheet.Name);
            xlsxBuilder.Append(entryStream, "\" sheetId=\""u8);
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, "\" r:id=\""u8);
            xlsxBuilder.Append(entryStream, worksheet.RelationshipId);
            xlsxBuilder.Append(entryStream, "\"></sheet>"u8);
        }

        xlsxBuilder.Append(entryStream,
            "</sheets>"u8
            + "</workbook>"u8);

        xlsxBuilder.Commit(entryStream);
    }

    private void AddWorkbookXmlRels()
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"u8
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"u8);

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<Relationship Id=\""u8);
            xlsxBuilder.Append(entryStream, worksheet.RelationshipId);
            xlsxBuilder.Append(entryStream, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"u8);
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, ".xml\" />"u8);
        }

        xlsxBuilder.Append(entryStream, "</Relationships>"u8);
        xlsxBuilder.Commit(entryStream);
    }

    private Worksheet BeginSheet(
        int id,
        string name,
        string relationshipId)
    {
        VerifyCanBeginSheet(id, name);

        // Make sure to end the previous sheet before beginning a new one.
        EndSheet();

        var entry = archive.CreateEntry($"xl/worksheets/sheet{id}.xml", compressionLevel);
        var entryStream = entry.Open();

        var worksheet = new Worksheet(
            xlsxBuilder,
            entryStream,
            this.stylesheet,
            id,
            name,
            relationshipId);
        worksheet.BeginSheet();
        worksheets.Add(worksheet);
        return worksheet;
    }

    private void EndSheet()
    {
        if (worksheets.Count == 0) return;

        worksheets[^1].EndSheet();
    }

    private void VerifyCanBeginSheet(
        int id,
        string name)
    {
        if (id < 0)
        {
            throw new InvalidOperationException("The XLSX format does not support negative identifiers.");
        }

        if (worksheets.Any(worksheet => worksheet.Id == id))
        {
            throw new InvalidOperationException($"A worksheet with identifier {id} was already added to the workbook.");
        }

        if (string.IsNullOrEmpty(name))
        {
            throw new InvalidOperationException("The XLSX format does not support an empty worksheet name.");
        }

        if (worksheets.Any(worksheet => worksheet.Name == name))
        {
            throw new InvalidOperationException($"A worksheet with name {name} was already added to the workbook.");
        }
    }
}
