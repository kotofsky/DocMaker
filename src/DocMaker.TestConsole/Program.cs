using DocMaker;
using DocMaker.TestConsole;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System.Diagnostics;
using System.IO.Packaging;
using System.Reflection;
using System.Text;



var countRowsElements = new int[] { 100, 1000, 10000, 100000 };
var excelTests = new ExcelTests();

for (int i = 0; i < countRowsElements.Length; i++)
{
    var stopWatch = new Stopwatch();
    stopWatch.Start();
    await excelTests.TestExcelDataAsync(countRowsElements[i]);
    stopWatch.Stop();

    Console.WriteLine($"Processed {countRowsElements[i]} rows for {stopWatch.Elapsed.ToString()}");
}









public class DataSet
{
    public List<DataRow> Rows { get; set; }
}

public class DataRow
{
    public int Index { get; set; }

    public Dictionary<string, string> Cells { get; set; }
}
