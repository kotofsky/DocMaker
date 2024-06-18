using DocMaker.Domain;
using System.Text;

namespace DocMaker.TestConsole;

internal class ExcelTests
{
    internal async Task TestExcelDataAsync(int countRows)
    {
        var FilePath = "TestData/Test_excel.xlsx";
        var OutputPath = "TestData/";
        var OutputFileName = "Result_excel.xlsx";

        var excelBuilder = new ExcelBuilder();

        var workFile = excelBuilder.CreateWorkFile(FilePath, OutputPath, OutputFileName);

        var data = new ExcelDataRows
        {
            Rows = new List<ExcelRow>(countRows)
        };

        string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        //100 - 214ms
        //1000 - 340ms
        //100000 - 7s146ms
        //1000000 - 56s900ms
        data.Rows.Add(new ExcelRow
        {
            RowIndex = 1,
            Cells = new Dictionary<string, string> { { "B1", "User" } }
        });

        data.Rows.Add(new ExcelRow
        {
            RowIndex = 2,
            Cells = new Dictionary<string, string> { { "B2", "City" } }
        });

        data.Rows.Add(new ExcelRow
        {
            RowIndex = 3,
            Cells = new Dictionary<string, string> { { "B3", "Email" } }
        });

        var builder = new StringBuilder();
        for (int i = 0; i < countRows; i++)
        {
            var row = new ExcelRow
            {
                RowIndex = i + 10,
                Cells = new Dictionary<string, string>(25)
            };

            for (int j = 0; j < 26; j++)
            {
                builder.Append(alphabet[j]);
                builder.Append(i + 10);
                row.Cells.Add(builder.ToString(), Guid.NewGuid().ToString());

                builder.Clear();
            }

            data.Rows.Add(row);
        }

        await excelBuilder.SetContentToFileAsync(workFile, "Sheet1", data);
    }
}
