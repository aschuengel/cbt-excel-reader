
using ClosedXML.Excel;
using ExcelReader;

const string docName = "Test.xlsx";
using var workbook = new XLWorkbook(docName);
foreach (var sheet in workbook.Worksheets)
{
    Console.WriteLine($"Processing sheet {sheet.Name}");
    var list = sheet.Map<Test>();
    var even = list.Where(item => item.IntValue % 2 == 0);
    foreach (var item in even)
    {
        Console.WriteLine($"{item.StringValue} {item.IntValue} {item.DoubleValue}");
    }
}