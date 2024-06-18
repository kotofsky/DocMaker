using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Reflection;
using System.IO.Packaging;
using DocMaker.Domain;

namespace DocMaker;

/// <inheritdoc />
public sealed class ExcelBuilder : IExcelBuilder
{
    /// <inheritdoc />
    public string CreateWorkFile(string templatePath, string outputPath, string outputFileName)
    {
        var appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        if (string.IsNullOrEmpty(appPath))
        {
            throw new ApplicationException("Base path of application not found");
        }

        var sourcePath = Path.Combine(appPath, templatePath);
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException($"Can't find template at this path {sourcePath}");
        }

        var resultPath = Path.Combine(appPath, outputPath);
        if (Directory.Exists(resultPath) == false)
        {
            Directory.CreateDirectory(resultPath);
        }

        var fullResultPath = Path.GetFullPath(resultPath + outputFileName);

        var buffer = new byte[1024 * 1024];

        using (FileStream sr = new(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (BufferedStream srb = new(sr))
        using (FileStream sw = new(fullResultPath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        using (BufferedStream swb = new(sw))
        {
            while (true)
            {
                var bytesRead = srb.Read(buffer, 0, buffer.Length);
                if (bytesRead == 0) break;
                swb.Write(buffer, 0, bytesRead);
            }
            swb.Flush();
        }

        return fullResultPath;
    }

    // <inheritdoc />
    public async Task SetContentToFileAsync(string filePath, string sheetName, ExcelDataRows contentData)
    {
        var worksheetPartId = await AddContentDataAsync(filePath, sheetName, contentData);
        await AddSheetToDocumentAsync(filePath, sheetName, worksheetPartId);
    }

    private async Task<string> AddContentDataAsync(string filePath, string sheetName, ExcelDataRows excelDataRows)
    {
        ExtractHeaders(filePath, sheetName, excelDataRows);

        await using var fileStream = File.Create(filePath);
        using var package = Package.Open(fileStream, FileMode.Create, FileAccess.Write);
        using var excel = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook);

        excel.AddWorkbookPart();

        var worksheetPart = excel.WorkbookPart.AddNewPart<WorksheetPart>();
        var worksheetPartId = excel.WorkbookPart.GetIdOfPart(worksheetPart);

        OpenXmlWriter openXmlWriter = OpenXmlWriter.Create(worksheetPart);

        openXmlWriter.WriteStartElement(new Worksheet());
        openXmlWriter.WriteStartElement(new SheetData());

        // write data rows
        foreach (var row in excelDataRows.Rows.OrderBy(r => r.RowIndex))
        {
            Row r = new()
            {
                RowIndex = (uint)row.RowIndex
            };

            openXmlWriter.WriteStartElement(r);

            foreach (var rowCell in row.Cells)
            {
                Cell c = new()
                {
                    DataType = CellValues.String,
                    CellReference = rowCell.Key
                };

                //cell start
                openXmlWriter.WriteStartElement(c);

                CellValue v = new(rowCell.Value);
                //cell value
                openXmlWriter.WriteElement(v);

                //cell end
                openXmlWriter.WriteEndElement();
            }

            // end row
            openXmlWriter.WriteEndElement();
        }

        // sheetdata end
        openXmlWriter.WriteEndElement();
        // worksheet end
        openXmlWriter.WriteEndElement();

        openXmlWriter.Close();

        return worksheetPartId;
    }

    private void ExtractHeaders(string filePath, string sheetName, ExcelDataRows excelDataRows)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            throw new FileNotFoundException($"Report file not found at this path: {filePath}");
        }

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
        {
            var wbPart = document.WorkbookPart;
            var theSheet = (wbPart?.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Value.Trim() == sheetName)) 
                ?? throw new ArgumentException($"Sheetname {sheetName} not found");
            
            WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id);
            Worksheet worksheet = wsPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            var rows = sheetData?.Elements<Row>().ToArray();

            for (int i = 0; i < rows.Length; i++)
            {
                var cells = rows[i].Elements<Cell>().Where(c => !string.IsNullOrWhiteSpace(c.InnerText));

                foreach (var cell in cells)
                {
                    var cellValue = cell.InnerText.Replace(" ", "");
                    if (cell.DataType is not null)
                    {
                        if (cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTable is not null)
                            {
                                cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
                            }
                        }
                        else if (cell.DataType.Value == CellValues.Boolean)
                        {
                            switch (cellValue)
                            {
                                case "0":
                                    cellValue = "FALSE";
                                    break;
                                default:
                                    cellValue = "TRUE";
                                    break;
                            }
                        }
                    }

                    cellValue = cellValue.Trim().Replace(" ", "");
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        int rowIndex = int.Parse(cell?.CellReference.Value[1..]);

                        var row = excelDataRows.Rows.FirstOrDefault(r => r.RowIndex == rowIndex);

                        if (row == null)
                        {
                            row = new ExcelRow
                            {
                                RowIndex = rowIndex,
                                Cells = new Dictionary<string, string> {
                                    { cell.CellReference.Value, cellValue }
                                }
                            };
                            excelDataRows.Rows.Add(row);
                        }
                        else
                        {
                            row.Cells.Add(cell.CellReference.Value, cellValue);
                            row.Cells = row.Cells.OrderBy(c => c.Key).ToDictionary();
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Finishes the process of writing data files.
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="worksheetPartId"></param>
    private async Task AddSheetToDocumentAsync(string filePath, string sheetName, string worksheetPartId)
    {
        await using var fileStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var package = Package.Open(fileStream, FileMode.Open, FileAccess.ReadWrite);
        using var excel = SpreadsheetDocument.Open(package);

        if (excel.WorkbookPart is null)
            throw new InvalidOperationException("Workbook part cannot be null!");

        var xmlWriter = OpenXmlWriter.Create(excel.WorkbookPart);
        xmlWriter.WriteStartElement(new Workbook());
        xmlWriter.WriteStartElement(new Sheets());

        xmlWriter.WriteElement(new Sheet { Id = worksheetPartId, Name = sheetName, SheetId = 1 });

        // sheets end
        xmlWriter.WriteEndElement();

        // workbook end
        xmlWriter.WriteEndElement();

        xmlWriter.Close();
        xmlWriter.Dispose();
    }
}