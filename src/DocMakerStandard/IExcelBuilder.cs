using DocMaker.Domain;

namespace DocMaker;

public interface IExcelBuilder
{
    string CreateWorkFile(string templatePath, string outputPath, string outputFileName);
    Task SetContentToFileAsync(string filePath, string sheetName, ExcelDataRows contentData);
}