# TinyXlsx

TinyXlsx is a lightweight and efficient library designed for writing Excel files in the XLSX format. It focuses on optimal performance by avoiding unnecessary overhead at all costs.

The library is built for .NET 8.0, ensuring compatibility with the latest versions of the framework. It supports two primary modes of writing data:

1. Writing to a `MemoryStream` for in-memory processing.
1. Writing to a `FileStream` to save the generated Excel file directly to disk.

TinyXlsx focuses on simplicity, providing only the necessary functionality to perform basic Excel file operations with minimal resource usage. Future versions may include more advanced features like reading and manipulating existing Excel files.

# Requirements

- .NET 8.0

# Features

Supported:
1. Writing to a `MemoryStream` for in-memory processing.
1. Writing to a `FileStream` to save the generated Excel file directly to disk.
1. Precise cell and row positioning.
    1. By default, `BeginRow` automatically progresses to the next row, and `WriteCellValue` automatically writes to the cell in the next column.
    1. An index can be specified using `BeginRowAt` and `WriteCellValueAt` if a row or column must be skipped.
1. Writing formulas.

Not supported yet:
1. Reading an existing document.
1. Editing an existing document.
1. Images.
1. Charts.
1. Cell merging.
1. Rich text.
1. Conditional formatting.
1. Comments.
1. Hyperlinks.

## Writing to a `MemoryStream`

By default the `Workbook` writes to a `MemoryStream`. This method should be used in scenarios where a file does not need to be stored locally but is instead intended to be sent directly to a client via a website or similar service.

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

for (var i = 0; i < 100; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(true);
    worksheet.WriteCellValue(0.1m);
    worksheet.WriteCellValue(123.456);
    worksheet.WriteCellValue(DateTime.Now);
    worksheet.WriteCellValue(DateTime.Now, "yyyy/MM/dd");
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

By supplying a `string` parameter to the `Workbook` constructor, the `Workbook` writes to a file.

```csharp
using TinyXlsx;

using var workbook = new Workbook("fileName.xlsx");
var worksheet = workbook.BeginSheet();

for (var i = 0; i < 100; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(true);
    worksheet.WriteCellValue(0.1m);
    worksheet.WriteCellValue(123.456);
    worksheet.WriteCellValue(DateTime.Now);
    worksheet.WriteCellValue(DateTime.Now, "yyyy/MM/dd");
    worksheet.WriteCellValue("Text");
    worksheet.WriteCellValue(123.456, "0.00");
    worksheet.WriteCellValue(123.456, "0.00%");
    worksheet.WriteCellValue(123.456, "0.00E+00");
    worksheet.WriteCellValue(123.456, "$#,##0.00");
    worksheet.WriteCellValue(123.456, "#,##0.00 [$USD]");
}
workbook.Close();
```

## Precise cell and row positioning

By default, `BeginRow` automatically progresses to the next row, and `WriteCellValue` automatically writes to the cell in the next column. An index can be specified using `BeginRowAt` and `WriteCellValueAt` if a row or column must be skipped. Going backwards is not supported due to the streaming nature of the library.

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

worksheet.BeginRowAt(10);
worksheet.WriteCellValueAt(5, 123.456);
worksheet.BeginRow(); // Begins row 11.
worksheet.WriteCellValue(DateTime.Now); // Writes in first cell on row 11.

var stream = workbook.Close();
```

## Writing formulas

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

var i = 0;
for (; i < 10; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(0.1m);
    worksheet.WriteCellValue(0.2m);
    worksheet.WriteCellValue(0.3m);
    worksheet.WriteCellFormula($"=SUM(A{i + 1}:C{i + 1})");
}
worksheet.BeginRow();
worksheet.WriteCellFormula($"=SUM(A1:A{i})");
worksheet.WriteCellFormula($"=SUM(B1:B{i})");
worksheet.WriteCellFormula($"=SUM(C1:C{i})");

var stream = workbook.Close();
```

# Optimization

For in-memory scenarios the default capacity is set to 64 KB. However, if the document size is known to be much larger in advance, it is recommended to set an initial capacity which more closely aligns with this size. An initial capacity can be given to the `Workbook` constructor. The default `MemoryStream` will automatically resize as data is written, but setting a capacity upfront reduces the overhead caused by repeated internal buffer expansions.

```csharp
using TinyXlsx;

var initialCapacity = 1024 * 1024; // 1 MB.
using var workbook = new Workbook(initialCapacity);
var worksheet = workbook.BeginSheet();

// Add data here...

var stream = workbook.Close();
```

# Benchmarks

## Writing to a `MemoryStream`

NA means the library does not support writing to a `MemoryStream`.

100 records, 8 columns:

| Method    | Mean            | Error         | StdDev       | Gen0        | Gen1        | Gen2       | Allocated      |
|---------- |----------------:|--------------:|-------------:|------------:|------------:|-----------:|---------------:|
| ClosedXML |      3,001.8 us |      17.30 us |     15.34 us |     78.1250 |           - |          - |     1360.37 KB |
| NPOI      |      3,825.3 us |      65.89 us |     55.02 us |    125.0000 |     31.2500 |          - |     2117.53 KB |
| OpenXML   |        871.9 us |       5.02 us |      4.69 us |     35.1563 |     19.5313 |          - |      621.33 KB |
| TinyXlsx  |        645.8 us |       4.72 us |      4.41 us |      3.9063 |           - |          - |       73.60 KB |

10,000 records, 8 columns:

| Method    | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated   |
|---------- |----------:|---------:|---------:|----------:|----------:|----------:|------------:|
| ClosedXML | 220.82 ms | 1.297 ms | 1.213 ms | 6000.0000 | 2000.0000 | 1000.0000 | 99992.92 KB |
| NPOI      | 100.73 ms | 1.397 ms | 1.166 ms | 3500.0000 | 1000.0000 |         - | 60048.23 KB |
| OpenXML   | 141.68 ms | 2.465 ms | 2.058 ms | 3333.3333 | 3000.0000 | 1000.0000 | 54245.12 KB |
| TinyXlsx  |  57.05 ms | 0.438 ms | 0.410 ms |  222.2222 |  222.2222 |  222.2222 |   970.52 KB |

1,000,000 records, 8 columns:

| Method    | Mean     | Error    | StdDev   | Gen0        | Gen1        | Gen2       | Allocated   |
|---------- |---------:|---------:|---------:|------------:|------------:|-----------:|------------:|
| ClosedXML | 26.922 s | 0.0738 s | 0.0691 s | 541000.0000 |  80000.0000 | 10000.0000 | 10087.41 MB |
| NPOI      |  9.925 s | 0.0736 s | 0.0652 s | 357000.0000 |  90000.0000 |  1000.0000 |  5748.21 MB |
| OpenXML   | 15.220 s | 0.2343 s | 0.2077 s | 263000.0000 | 262000.0000 |  8000.0000 |  4858.12 MB |
| TinyXlsx  |  5.448 s | 0.0092 s | 0.0081 s |   2000.0000 |   2000.0000 |  2000.0000 |    63.99 MB |