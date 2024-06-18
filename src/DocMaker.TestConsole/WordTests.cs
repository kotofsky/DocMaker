using System.Reflection;

namespace DocMaker.TestConsole;

internal class WordTests
{
    internal async Task TestWordDataAsync()
    {
        var FilePath = "TestData/Test_word.docx";
        var OutputPath = "TestData/";
        var OutputFileName = "Result_word.docx";

        var wordBuilder = new WordBuilder();

        var template = wordBuilder.CreateTemplate();

        template.FieldsCollection.Add("test_data", "OMG! THIS IS THE TEST DATA");

        var docTables = new List<Domain.DocTable>();

        var firstTable = new Domain.DocTable();
        firstTable.AddRow(["FirstTableData1", "FirstTableData2", "FirstTableData3", "FirstTableData4", "FirstTableData5"]);
        docTables.Add(firstTable);

        var secondTable = new Domain.DocTable();
        secondTable.AddRow(["SecondTableData1", "-", "SecondTableData3"]);
        docTables.Add(secondTable);

        template.Tables = docTables.ToArray();

        var appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var templatePath = Path.Combine(appPath, FilePath);
        var resultPath = Path.Combine(appPath, OutputPath, OutputFileName);

        using (var stream = File.OpenRead(templatePath))
        {
            var result = await wordBuilder.BuildAsync(stream, template);

            using (Stream streamToWriteTo = File.Open(resultPath, FileMode.Create))
            {
                result.Seek(0, SeekOrigin.Begin);
                await result.CopyToAsync(streamToWriteTo);
            }
        }
    }
}
