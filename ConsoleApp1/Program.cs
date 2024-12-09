﻿using TinyXlsx;

using var workbook = new Workbook();
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
    worksheet.EndRow();
}
workbook.EndSheet();
var stream = workbook.Close();


using var fileStream = File.Create("smallest.xlsx");
stream.CopyTo(fileStream);
await fileStream.FlushAsync();
fileStream.Close();

//Console.WriteLine(System.Diagnostics.Process.GetCurrentProcess().PrivateMemorySize64);

//Console.ReadLine();
