using TinyXlsx;

using var workbook = new Workbook();
var worksheet = workbook.BeginSheet();

for (var i = 0; i < 10; i++)
{
    worksheet.BeginRow(i);
    worksheet.WriteCellValueAt(0, 123.456);
    worksheet.WriteCellValueAt(1, DateTime.Now);
    worksheet.WriteCellValueAt(2, "Text");
    worksheet.WriteCellValueAt(3, 123.456, "0.00");
    worksheet.WriteCellValueAt(4, 123.456, "0.00%");
    worksheet.WriteCellValueAt(5, 123.456, "0.00E+00");
    worksheet.WriteCellValueAt(6, 123.456, "$#,##0.00");
    worksheet.WriteCellValueAt(7, 123.456, "#,##0.00 [$USD]");
}

var worksheet2 = workbook.BeginSheet();

for (var i = 0; i < 10; i++)
{
    worksheet2.BeginRow(i);
    worksheet2.WriteCellValueAt(0, 123.456);
    worksheet2.WriteCellValueAt(1, DateTime.Now);
    worksheet2.WriteCellValueAt(2, "Text");
    worksheet2.WriteCellValueAt(3, 123.456, "0.00");
    worksheet2.WriteCellValueAt(4, 123.456, "0.00%");
    worksheet2.WriteCellValueAt(5, 123.456, "0.00E+00");
    worksheet2.WriteCellValueAt(6, 123.456, "$#,##0.00");
    worksheet2.WriteCellValueAt(7, 123.456, "#,##0.00 [$USD]");
}
var stream = workbook.Close();


using var fileStream = File.Create("smallest.xlsx");
stream.CopyTo(fileStream);
await fileStream.FlushAsync();
fileStream.Close();

//Console.WriteLine(System.Diagnostics.Process.GetCurrentProcess().PrivateMemorySize64);

//Console.ReadLine();
