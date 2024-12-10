# TinyXlsx
TinyXlsx is a lightweight and efficient library designed for writing Excel files in the XLSX format. It focuses on optimal performance by avoiding unnecessary overhead at all costs.

The library is built for .NET 8.0, ensuring compatibility with the latest versions of the framework. It supports two primary modes of writing data:

1.  Writing to a `MemoryStream` for in-memory processing.
2.  Writing to a `FileStream` to save the generated Excel file directly to disk.

TinyXlsx focuses on simplicity, providing only the necessary functionality to perform basic Excel file operations with minimal resource usage. Future versions may include more advanced features like reading and manipulating existing Excel files.

# Requirements
- .NET 8.0

# Features
Reading not supported yet.

1.  Writing to a `MemoryStream` for in-memory processing.
2.  Writing to a `FileStream` to save the generated Excel file directly to disk.

## Writing to a `MemoryStream`

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

for (var i = 0; i < 10_000; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(123.456);
    worksheet.WriteCellValue(DateTime.Now);
    worksheet.WriteCellValue("Text");
    worksheet.WriteCellValue(123.456, "0.00");
    worksheet.WriteCellValue(123.456, "0.00%");
    worksheet.WriteCellValue(123.456, "0.00E+00");
    worksheet.WriteCellValue(123.456, "$#,##0.00");
    worksheet.WriteCellValue(123.456, "#,##0.00 [$USD]");
}
var stream = workbook.Close();
```

## Writing to a `FileStream`

```csharp
using TinyXlsx;

using var workbook = new Workbook("fileName.xlsx");
var worksheet = workbook.BeginSheet();

for (var i = 0; i < 10_000; i++)
{
    worksheet.BeginRow(i);
    worksheet.WriteCellValue(0, 123.456);
    worksheet.WriteCellValue(1, DateTime.Now);
    worksheet.WriteCellValue(2, "Text");
    worksheet.WriteCellValue(3, 123.456, "0.00");
    worksheet.WriteCellValue(4, 123.456, "0.00%");
    worksheet.WriteCellValue(5, 123.456, "0.00E+00");
    worksheet.WriteCellValue(6, 123.456, "$#,##0.00");
    worksheet.WriteCellValue(7, 123.456, "#,##0.00 [$USD]");
}
workbook.Close();
```
# Benchmarks
| Method    | Records | Mean            | Error         | StdDev        | Gen0        | Gen1        | Gen2       | Allocated      |
|---------- |-------- |----------------:|--------------:|--------------:|------------:|------------:|-----------:|---------------:|
| ClosedXML | 100     |      3,025.4 us |      16.90 us |      15.81 us |     78.1250 |           - |          - |     1360.38 KB |
| NPOI      | 100     |      3,932.1 us |      77.64 us |     111.35 us |    125.0000 |     31.2500 |          - |     2117.36 KB |
| OpenXML   | 100     |        880.6 us |       3.59 us |       3.00 us |     31.2500 |     15.6250 |          - |      621.33 KB |
| TinyXlsx  | 100     |        720.3 us |       4.08 us |       3.82 us |    333.0078 |    333.0078 |   333.0078 |     1033.76 KB |
| ClosedXML | 10000   |    222,782.9 us |   1,175.39 us |   1,041.96 us |   6000.0000 |   2000.0000 |  1000.0000 |    99994.04 KB |
| NPOI      | 10000   |     99,148.6 us |     869.82 us |     726.34 us |   3500.0000 |   1000.0000 |          - |    60048.81 KB |
| OpenXML   | 10000   |    140,493.0 us |   2,774.45 us |   3,407.28 us |   3333.3333 |   3000.0000 |  1000.0000 |    54245.13 KB |
| TinyXlsx  | 10000   |     55,932.6 us |     654.49 us |     612.21 us |    300.0000 |    300.0000 |   300.0000 |     1034.36 KB |
| ClosedXML | 1000000 | 26,812,815.0 us | 101,568.95 us |  90,038.23 us | 541000.0000 |  80000.0000 | 10000.0000 | 10329501.95 KB |
| NPOI      | 1000000 |  9,782,638.5 us |  73,017.24 us |  64,727.88 us | 357000.0000 |  90000.0000 |  1000.0000 |  5886191.24 KB |
| OpenXML   | 1000000 | 14,563,814.5 us | 184,663.48 us | 172,734.33 us | 263000.0000 | 262000.0000 |  8000.0000 |  4974719.98 KB |
| TinyXlsx  | 1000000 |  5,555,489.6 us |  48,403.93 us |  45,277.06 us |   2000.0000 |   2000.0000 |  2000.0000 |     64551.7 KB |