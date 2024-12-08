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
var memoryStream = await workbook.BeginStreamAsync();
using var worksheet = await workbook.BeginSheetAsync();

for (var i = 0; i < 10; i++)
{
    await worksheet.BeginRowAsync(i);
    await worksheet.WriteCellValueAsync(0, 123.456);
    await worksheet.WriteCellValueAsync(1, DateTime.Now);
    await worksheet.WriteCellValueAsync(2, "Text");
    await worksheet.WriteCellValueAsync(3, 123.456, "0.00");
    await worksheet.WriteCellValueAsync(4, 123.456, "0.00%");
    await worksheet.WriteCellValueAsync(5, 123.456, "0.00E+00");
    await worksheet.WriteCellValueAsync(6, 123.456, "$#,##0.00");
    await worksheet.WriteCellValueAsync(7, 123.456, "#,##0.00 [$USD]");
    await worksheet.EndRowAsync();
}
await workbook.EndSheetAsync();
await workbook.EndStreamAsync();
```

## Writing to a `FileStream`

```csharp
using TinyXlsx;

using var workbook = new Workbook();
await workbook.BeginFileAsync("fileName.xlsx");
using var worksheet = await workbook.BeginSheetAsync();

for (var i = 0; i < 10; i++)
{
    await worksheet.BeginRowAsync(i);
    await worksheet.WriteCellValueAsync(0, 123.456);
    await worksheet.WriteCellValueAsync(1, DateTime.Now);
    await worksheet.WriteCellValueAsync(2, "Text");
    await worksheet.WriteCellValueAsync(3, 123.456, "0.00");
    await worksheet.WriteCellValueAsync(4, 123.456, "0.00%");
    await worksheet.WriteCellValueAsync(5, 123.456, "0.00E+00");
    await worksheet.WriteCellValueAsync(6, 123.456, "$#,##0.00");
    await worksheet.WriteCellValueAsync(7, 123.456, "#,##0.00 [$USD]");
    await worksheet.EndRowAsync();
}
await workbook.EndSheetAsync();
await workbook.EndFileAsync();
```
# Benchmarks
| Method    | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2      | Allocated |
|---------- |----------:|---------:|---------:|----------:|----------:|----------:|----------:|
| ClosedXml | 222.31 ms | 2.202 ms | 2.060 ms | 6000.0000 | 2000.0000 | 1000.0000 |  97.65 MB |
| Npoi      |  99.61 ms | 0.925 ms | 0.722 ms | 3500.0000 | 1000.0000 |         - |  58.64 MB |
| OpenXml   | 137.23 ms | 2.699 ms | 3.603 ms | 3333.3333 | 3000.0000 | 1000.0000 |  52.97 MB |
| TinyXlsx  | 100.97 ms | 0.476 ms | 0.422 ms |  200.0000 |  200.0000 |  200.0000 |   1.01 MB |