﻿using System.IO.Compression;

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
    private readonly Dictionary<string, (int ZeroBasedIndex, int CustomFormatIndex)> numberFormats;
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
        numberFormats = [];
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
        numberFormats = [];
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
    /// Gets or creates a unique number format style for the specified format string.
    /// </summary>
    /// <param name="format">
    /// The format string to get or create.
    /// </param>
    /// <returns>
    /// A tuple containing the zero-based index and custom format index for the style.
    /// </returns>
    /// <exception cref="NotSupportedException">
    /// Thrown if the number of styles exceeds the maximum supported by the XLSX format.
    /// </exception>
    public (int ZeroBasedIndex, int CustomFormatIndex) GetOrCreateNumberFormat(string format)
    {
        var count = numberFormats.Count;

        if (count >= Constants.MaximumStyles)
        {
            throw new NotSupportedException("The XLSX format does not support more than 65,490 styles.");
        }

        if (numberFormats.TryGetValue(format, out var indexes))
        {
            return indexes;
        }

        indexes = (count + 1, count + 164);
        numberFormats.Add(format, indexes);
        return indexes;
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
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
            + "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />"
            + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
            + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
            + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<Override PartName=\"/xl/worksheets/sheet");
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />");
        }
        xlsxBuilder.Append(entryStream, "</Types>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddDocPropsAppXml()
    {
        var entry = archive.CreateEntry("docProps/app.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
            + "<Application>TinyXlsx 0.1.0</Application>"
            + "<AppVersion>15.0000</AppVersion>"
            + "</Properties>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddDocPropsCoreXml()
    {
        var entry = archive.CreateEntry("docProps/core.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<coreProperties "
            + "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
            + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
            + "xmlns:dcterms=\"http://purl.org/dc/terms/\""
            + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\""
            + "xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">"
            + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">");

        xlsxBuilder.Append(entryStream, DateTime.UtcNow.ToString("yyyy-MM-ddThh:mm:ssZ"));
        xlsxBuilder.Append(entryStream, "</dcterms:created><dc:creator></dc:creator></coreProperties>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddRels()
    {
        var entry = archive.CreateEntry("_rels/.rels", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\" />"
            + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
            + "</Relationships>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddSharedStringsXml()
    {
        var entry = archive.CreateEntry("xl/sharedStrings.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>"
            + "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "</sst>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddStylesXml()
    {
        var entry = archive.CreateEntry("xl/styles.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<numFmts count=\"");

        xlsxBuilder.Append(entryStream, numberFormats.Count);
        xlsxBuilder.Append(entryStream, "\">");
        foreach (var item in numberFormats)
        {
            xlsxBuilder.Append(entryStream, "<numFmt numFmtId=\"");
            xlsxBuilder.Append(entryStream, item.Value.CustomFormatIndex);
            xlsxBuilder.Append(entryStream, "\" formatCode=\"");
            xlsxBuilder.Append(entryStream, item.Key);
            xlsxBuilder.Append(entryStream, "\"/>");
        }

        xlsxBuilder.Append(entryStream,
            "</numFmts>"
            + "<fonts count=\"1\">"
            + "<font><sz val=\"11\"/><color indexed=\"8\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"
            + "</fonts>"
            + "<fills count=\"2\">"
            + "<fill><patternFill patternType=\"none\"/></fill>"
            + "<fill><patternFill patternType=\"darkGray\"/></fill>"
            + "</fills>"
            + "<borders count=\"1\">"
            + "<border><left/><right/><top/><bottom/><diagonal/></border>"
            + "</borders>"
            + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
            + "<cellXfs count=\"");

        xlsxBuilder.Append(entryStream, numberFormats.Count + 1);
        xlsxBuilder.Append(entryStream, "\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        foreach (var item in numberFormats)
        {
            xlsxBuilder.Append(entryStream, "<xf numFmtId=\"");
            xlsxBuilder.Append(entryStream, item.Value.CustomFormatIndex);
            xlsxBuilder.Append(entryStream, "\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
        }

        xlsxBuilder.Append(entryStream,
            "</cellXfs>"
            + "</styleSheet>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddWorkbookXml()
    {
        var entry = archive.CreateEntry("xl/workbook.xml", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<sheets>");

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<sheet name=\"");
            xlsxBuilder.Append(entryStream, worksheet.Name);
            xlsxBuilder.Append(entryStream, "\" sheetId=\"");
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, "\" r:id=\"");
            xlsxBuilder.Append(entryStream, worksheet.RelationshipId);
            xlsxBuilder.Append(entryStream, "\"></sheet>");
        }

        xlsxBuilder.Append(entryStream,
            "</sheets>"
            + "</workbook>");

        xlsxBuilder.Commit(entryStream);
    }

    private void AddWorkbookXmlRels()
    {
        var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", compressionLevel);
        using var entryStream = entry.Open();

        xlsxBuilder.Append(entryStream,
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");

        foreach (var worksheet in worksheets)
        {
            xlsxBuilder.Append(entryStream, "<Relationship Id=\"");
            xlsxBuilder.Append(entryStream, worksheet.RelationshipId);
            xlsxBuilder.Append(entryStream, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
            xlsxBuilder.Append(entryStream, worksheet.Id);
            xlsxBuilder.Append(entryStream, ".xml\" />");
        }

        xlsxBuilder.Append(entryStream, "</Relationships>");
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
            this,
            xlsxBuilder,
            entryStream,
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
