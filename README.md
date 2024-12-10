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
| Method    | Records | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|---------- |-------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| ClosedXml | 10000   | 225.90 ms | 3.009 ms | 2.814 ms | 6000.0000 | 2000.0000 | 1000.0000 |  97.65 MB |
| Npoi      | 10000   | 102.85 ms | 1.900 ms | 1.684 ms | 3500.0000 | 1000.0000 |         - |  58.64 MB |
| OpenXml   | 10000   | 142.64 ms | 2.601 ms | 3.290 ms | 3333.3333 | 3000.0000 | 1000.0000 |  52.97 MB |
| TinyXlsx  | 10000   |  55.49 ms | 0.434 ms | 0.406 ms |  222.2222 |  222.2222 |  222.2222 |   1.01 MB |