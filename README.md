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

for (var i = 1; i <= 100; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(true);
    worksheet.WriteCellValue(123456);
    worksheet.WriteCellValue(123.456m);
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

for (var i = 1; i <= 100; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(true);
    worksheet.WriteCellValue(123456);
    worksheet.WriteCellValue(123.456m);
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

By default, `BeginRow` automatically progresses to the next row, and `WriteCellValue` automatically writes to the cell in the next column. A one-based index can be specified using `BeginRowAt` and `WriteCellValueAt`, if for example a row or column must be skipped. Going backwards is not supported due to the streaming nature of the library.

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

var i = 1;
for (; i <= 10; i++)
{
    worksheet.BeginRow();
    worksheet.WriteCellValue(0.1m);
    worksheet.WriteCellValue(0.2m);
    worksheet.WriteCellValue(0.3m);
    worksheet.WriteCellFormula($"=SUM(A{i}:C{i})");
}
i++;
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

100 records, 12 columns:

| Method    | Mean       | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated   |
|---------- |-----------:|---------:|---------:|----------:|----------:|----------:|------------:|
| FastExcel |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| NanoXLSX  |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| PicoXLSX  |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| ClosedXML | 4,084.3 us | 34.52 us | 28.83 us |  109.3750 |   31.2500 |   15.6250 |  1894.83 KB |
| LargeXlsx |   666.4 us |  8.45 us |  7.49 us |  142.5781 |   83.9844 |         - |  2357.43 KB |
| MiniExcel | 2,385.8 us | 44.15 us | 41.30 us | 1683.5938 | 1667.9688 | 1667.9688 | 18309.52 KB |
| NPOI      | 4,301.8 us | 64.31 us | 57.01 us |  140.6250 |   46.8750 |         - |  2366.37 KB |
| OpenXML   | 1,308.5 us |  7.80 us |  7.30 us |   54.6875 |   46.8750 |   23.4375 |   911.62 KB |
| TinyXlsx  |   325.2 us |  2.53 us |  2.36 us |    4.8828 |    0.9766 |         - |    81.66 KB |

10,000 records, 12 columns:

| Method    | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|---------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| FastExcel |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| NanoXLSX  |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| PicoXLSX  |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| ClosedXML | 317.52 ms | 2.624 ms | 2.454 ms | 9000.0000 | 3000.0000 | 2000.0000 | 140.08 MB |
| LargeXlsx |  56.82 ms | 0.253 ms | 0.225 ms |  444.4444 |  333.3333 |  222.2222 |   8.75 MB |
| MiniExcel |  35.36 ms | 0.699 ms | 1.633 ms | 3538.4615 | 2923.0769 | 2153.8462 |  42.85 MB |
| NPOI      | 142.51 ms | 1.847 ms | 1.637 ms | 5000.0000 | 1000.0000 |         - |   81.2 MB |
| OpenXML   | 223.78 ms | 4.219 ms | 4.332 ms | 5333.3333 | 5000.0000 | 1666.6667 |  71.89 MB |
| TinyXlsx  |  22.81 ms | 0.070 ms | 0.066 ms | 468.7500  | 468.7500  | 468.7500  |   1.96 MB |

1,000,000 records, 12 columns:

| Method    | Mean     | Error    | StdDev   | Gen0        | Gen1        | Gen2       | Allocated   |
|---------- |---------:|---------:|---------:|------------:|------------:|-----------:|------------:|
| FastExcel |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| NanoXLSX  |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| PicoXLSX  |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| ClosedXML | 36.581 s | 0.2513 s | 0.2351 s | 802000.0000 | 100000.0000 | 10000.0000 | 14268.52 MB |
| LargeXlsx |  5.562 s | 0.0224 s | 0.0209 s |  41000.0000 |   7000.0000 |  6000.0000 |   691.31 MB |
| MiniExcel |  2.534 s | 0.0308 s | 0.0288 s | 166000.0000 |  32000.0000 |  7000.0000 |  2648.58 MB |
| NPOI      | 14.156 s | 0.1672 s | 0.1564 s | 500000.0000 | 167000.0000 |  1000.0000 |   8098.2 MB |
| OpenXML   | 22.693 s | 0.2360 s | 0.2208 s | 383000.0000 | 382000.0000 |  9000.0000 |  7518.95 MB |
| TinyXlsx  |  2.221 s | 0.0086 s | 0.0072 s |   3000.0000 |   3000.0000 |  3000.0000 |   127.97 MB |