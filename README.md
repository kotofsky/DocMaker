# DocMaker

A versatile library for generating Word documents with content data and creating tables is crucial for many applications. This library not only supports the creation of detailed Word documents but also enables generating Excel files using predefined templates, even with very large datasets, ensuring consistency and efficiency in document automation.

[![NuGet](https://img.shields.io/nuget/v/DocMaker.svg)](https://www.nuget.org/packages/DocMaker/)
[![Downloads](https://img.shields.io/nuget/dt/DocMaker.svg)](https://www.nuget.org/packages/DocMaker/)
![Build status](https://github.com/kotofsky/DocMaker/actions/workflows/main.yml/badge.svg)


### Adding content data to Word document
```
var wordBuilder = new WordBuilder();
var template = wordBuilder.CreateTemplate();

//key - name of content data wrapper in your Word document, value - raw data
template.FieldsCollection.Add("test_data", "OMG! THIS IS THE TEST DATA");

//here stream is the Stream of your template
var result = await wordBuilder.BuildAsync(stream, template);
```
### Adding tables to Word document
```
var wordBuilder = new WordBuilder();
var template = wordBuilder.CreateTemplate();

var docTables = new List<DocTable>();

//each table object contains rows and cells.
var firstTable = new DocTable();

//you can pass symbol '-' to merge cells in your table
firstTable.AddRow(["FirstTableData1", "FirstTableData2", "FirstTableData3", 
                                        "FirstTableData4", "FirstTableData5"]);
docTables.Add(firstTable);
template.Tables = docTables.ToArray();

//here stream is the Stream of your template
//BuildAsync method will find any tables in your Word template(it should contains headers or something)
//and will add new rows and cells
var result = await wordBuilder.BuildAsync(stream, template);
```

### Generate Excel files
```
var excelBuilder = new ExcelBuilder();
//create copy of your template
var workFile = excelBuilder.CreateWorkFile(FilePath, OutputPath, OutputFileName);

//create object with all data
var data = new ExcelDataRows
{
    Rows = new List<ExcelRow>(countRows)
};

//add some data
data.Rows.Add(new ExcelRow
{
    RowIndex = 1,
    Cells = new Dictionary<string, string> { { "B1", "User" } }
});

//fill data to your result file. you can pass any sheet name that you need to process
await excelBuilder.SetContentToFileAsync(workFile, "Sheet1", data);
```

## License

The DocMaker is licensed under the MIT license.