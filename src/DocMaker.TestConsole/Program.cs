using DocMaker.TestConsole;
using System.Diagnostics;

//Excel
var countRowsElements = new int[] { 100, 1000, 10000, 100000 };
var excelTests = new ExcelTests();
var stopWatch = new Stopwatch();

for (int i = 0; i < countRowsElements.Length; i++)
{
    stopWatch.Start();
    await excelTests.TestExcelDataAsync(countRowsElements[i]);
    stopWatch.Stop();
    Console.WriteLine($"Processed {countRowsElements[i]} rows for {stopWatch.Elapsed.ToString()}");
    stopWatch.Reset();
}


//Word
var wordTests = new  WordTests();

stopWatch.Start();
await wordTests.TestWordDataAsync();
stopWatch.Stop();
Console.WriteLine($"Processed word file generation with {stopWatch.Elapsed.ToString()}");