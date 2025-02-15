# TinyXlsx

TinyXlsx is a lightweight and efficient library designed for writing Excel files in the XLSX format. It focuses on optimal performance by avoiding unnecessary overhead at all costs.

1. [About](#about)
1. [Benchmarks](#benchmarks)
1. [Requirements](#requirements)
1. [Features](#features)
1. [Optimization](#optimization)

## About

[![NuGet Version](https://img.shields.io/nuget/v/TinyXlsx?style=for-the-badge)](https://www.nuget.org/packages/TinyXlsx)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TinyXlsx?style=for-the-badge)](https://www.nuget.org/packages/TinyXlsx)

The library is built for .NET 8.0, ensuring compatibility with the latest versions of the framework. It supports two primary modes of writing data:

1. Writing to a `MemoryStream` for in-memory processing.
1. Writing to a `FileStream` to save the generated XLSX file directly to disk.

TinyXlsx focuses on simplicity, providing only the necessary functionality to perform basic XLSX file operations with minimal resource usage. Future versions may include more advanced features like reading and manipulating existing XLSX files.

## Benchmarks

### Writing to a `MemoryStream`

NA means the library does not support writing to a `MemoryStream`.

100 records, 12 columns:

| Method    | Mean       | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated   |
|---------- |-----------:|---------:|---------:|----------:|----------:|----------:|------------:|
| FastExcel |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| NanoXLSX  |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| PicoXLSX  |         NA |       NA |       NA |        NA |        NA |        NA |          NA |
| ClosedXML | 4,084.3 us | 34.52 us | 28.83 us |  109.3750 |   31.2500 |   15.6250 |  1894.83 KB |
| LargeXlsx |   666.4 us |  8.45 us |  7.49 us |  142.5781 |   83.9844 |         - |  2357.43 KB |
| MiniExcel | 4,333.0 us | 68.15 us | 63.75 us | 1632.8125 | 1617.1875 | 1617.1875 | 18637.51 KB |
| NPOI      | 4,301.8 us | 64.31 us | 57.01 us |  140.6250 |   46.8750 |         - |  2366.37 KB |
| OpenXML   | 1,308.5 us |  7.80 us |  7.30 us |   54.6875 |   46.8750 |   23.4375 |   911.62 KB |
| TinyXlsx  |   331.7 us |  1.04 us |  0.92 us |    4.8828 |    0.9766 |         - |    81.66 KB |

10,000 records, 12 columns:

| Method    | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|---------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| FastExcel |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| NanoXLSX  |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| PicoXLSX  |        NA |       NA |       NA |        NA |        NA |        NA |        NA |
| ClosedXML | 317.52 ms | 2.624 ms | 2.454 ms | 9000.0000 | 3000.0000 | 2000.0000 | 140.08 MB |
| LargeXlsx |  56.82 ms | 0.253 ms | 0.225 ms |  444.4444 |  333.3333 |  222.2222 |   8.75 MB |
| MiniExcel |  57.42 ms | 1.119 ms | 1.569 ms | 5555.5556 | 3555.5556 | 2666.6667 |  68.07 MB |
| NPOI      | 142.51 ms | 1.847 ms | 1.637 ms | 5000.0000 | 1000.0000 |         - |   81.2 MB |
| OpenXML   | 223.78 ms | 4.219 ms | 4.332 ms | 5333.3333 | 5000.0000 | 1666.6667 |  71.89 MB |
| TinyXlsx  |  23.24 ms | 0.120 ms | 0.112 ms | 468.7500  | 468.7500  | 468.7500  |   1.96 MB |

1,000,000 records, 12 columns:

| Method    | Mean     | Error    | StdDev   | Gen0        | Gen1        | Gen2       | Allocated   |
|---------- |---------:|---------:|---------:|------------:|------------:|-----------:|------------:|
| FastExcel |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| NanoXLSX  |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| PicoXLSX  |       NA |       NA |       NA |          NA |          NA |         NA |          NA |
| ClosedXML | 36.581 s | 0.2513 s | 0.2351 s | 802000.0000 | 100000.0000 | 10000.0000 | 14268.52 MB |
| LargeXlsx |  5.562 s | 0.0224 s | 0.0209 s |  41000.0000 |   7000.0000 |  6000.0000 |   691.31 MB |
| MiniExcel |  5.258 s | 0.0282 s | 0.0263 s | 322000.0000 |  14000.0000 |  7000.0000 |  5215.86 MB |
| NPOI      | 14.156 s | 0.1672 s | 0.1564 s | 500000.0000 | 167000.0000 |  1000.0000 |   8098.2 MB |
| OpenXML   | 22.693 s | 0.2360 s | 0.2208 s | 383000.0000 | 382000.0000 |  9000.0000 |  7518.95 MB |
| TinyXlsx  |  2.290 s | 0.0103 s | 0.0086 s |   3000.0000 |   3000.0000 |  3000.0000 |   127.97 MB |

## Requirements

- .NET 8.0

## Features

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

### Writing to a `MemoryStream`

By default the `Workbook` writes to a `MemoryStream`. This method should be used in scenarios where a file does not need to be stored locally but is instead intended to be sent directly to a client via a website or similar service.

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

for (var i = 1; i <= 100; i++)
{
    worksheet
        .BeginRow()
        .WriteCellValue(true)
        .WriteCellValue(123456)
        .WriteCellValue(123.456m)
        .WriteCellValue(123.456)
        .WriteCellValue(DateTime.Now)
        .WriteCellValue(DateTime.Now, "yyyy/MM/dd")
        .WriteCellValue("Text")
        .WriteCellValue(123.456, "0.00")
        .WriteCellValue(123.456, "0.00%")
        .WriteCellValue(123.456, "0.00E+00")
        .WriteCellValue(123.456, "$#,##0.00")
        .WriteCellValue(123.456, "#,##0.00 [$USD]");
}
var stream = workbook.Close();
```

### Writing to a `FileStream`

By supplying a `string` parameter to the `Workbook` constructor, the `Workbook` writes to a file.

```csharp
using TinyXlsx;

using var workbook = new Workbook("fileName.xlsx");
var worksheet = workbook.BeginSheet();

for (var i = 1; i <= 100; i++)
{
    worksheet
        .BeginRow()
        .WriteCellValue(true)
        .WriteCellValue(123456)
        .WriteCellValue(123.456m)
        .WriteCellValue(123.456)
        .WriteCellValue(DateTime.Now)
        .WriteCellValue(DateTime.Now, "yyyy/MM/dd")
        .WriteCellValue("Text")
        .WriteCellValue(123.456, "0.00")
        .WriteCellValue(123.456, "0.00%")
        .WriteCellValue(123.456, "0.00E+00")
        .WriteCellValue(123.456, "$#,##0.00")
        .WriteCellValue(123.456, "#,##0.00 [$USD]");
}
workbook.Close();
```

### Precise cell and row positioning

By default, `BeginRow` automatically progresses to the next row, and `WriteCellValue` automatically writes to the cell in the next column. A one-based index can be specified using `BeginRowAt` and `WriteCellValueAt`, if for example a row or column must be skipped. Going backwards is not supported due to the streaming nature of the library.

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

worksheet.BeginRowAt(10); // Begins row 10.

worksheet.WriteCellValueAt(5, 123.456); // Writes in the fifth cell on row 10.

worksheet.BeginRow(); // Begins row 11.

worksheet.WriteCellValue(DateTime.Now); // Writes in the first cell on row 11.

var stream = workbook.Close();
```

### Writing formulas

Formulas can be added using `WriteCellFormula`. The library does not validate or calculate any formula, it is written as-is into the cell.

```csharp
using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

var i = 1;
for (; i <= 10; i++)
{
    worksheet
        .BeginRow()
        .WriteCellValue(0.1m)
        .WriteCellValue(0.2m)
        .WriteCellValue(0.3m)
        .WriteCellFormula($"=SUM(A{i}:C{i})");
}
i++;
worksheet
    .BeginRow()
    .WriteCellFormula($"=SUM(A1:A{i})")
    .WriteCellFormula($"=SUM(B1:B{i})")
    .WriteCellFormula($"=SUM(C1:C{i})");

var stream = workbook.Close();
```

## Optimization

For in-memory scenarios the default capacity is set to 64 KB. However, if the document size is known to be much larger in advance, it is recommended to set an initial capacity which more closely aligns with this size. An initial capacity can be given to the `Workbook` constructor. The default `MemoryStream` will automatically resize as data is written, but setting a capacity upfront reduces the overhead caused by repeated internal buffer expansions.

```csharp
using TinyXlsx;

var initialCapacity = 1024 * 1024 * 64; // 64 MB.
using var workbook = new Workbook(initialCapacity);
var worksheet = workbook.BeginSheet();

// Add data here...

var stream = workbook.Close();
```
