using TinyXlsx;

using var workbook = new Workbook();
var stream = await workbook.BeginStreamAsync(4 * 1024 * 1024);
var worksheet = await workbook.BeginSheetAsync();

for (var i = 0; i < 100_000; i++)
{
    await worksheet.BeginRowAsync(i);
    worksheet.WriteCellValue(0, 123.456);
    worksheet.WriteCellValue(1, DateTime.Now);
    worksheet.WriteCellValue(2, "Text");
    worksheet.WriteCellValue(3, 123.456, "0.00");
    worksheet.WriteCellValue(4, 123.456, "0.00%");
    worksheet.WriteCellValue(5, 123.456, "0.00E+00");
    worksheet.WriteCellValue(6, 123.456, "$#,##0.00");
    worksheet.WriteCellValue(7, 123.456, "#,##0.00 [$USD]");
    await worksheet.EndRowAsync();
}
await workbook.EndSheetAsync();
await workbook.EndStreamAsync();

//using var fileStream = File.Create("123456_optimal.xlsx");
//stream.CopyTo(fileStream);
//await fileStream.FlushAsync();
//fileStream.Close();

Console.WriteLine(System.Diagnostics.Process.GetCurrentProcess().PrivateMemorySize64);

Console.ReadLine();
