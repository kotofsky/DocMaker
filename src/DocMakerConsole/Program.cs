using System.Collections.Generic;
using System.IO;
using DocMaker;
using DocMaker.Domain;

namespace DocMakerConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = "1.docx";

            var maker = new WordBuilder();
            var template = maker.CreateTemplate();
            template.FieldsCollection = new Dictionary<string, string>();
            template.FieldsCollection.Add("test","2333");

            var table1 = new DocTable();
            var cells = new[] {"a", "b", "c", "d"};
            table1.AddRow(cells);
            template.Tables = new[] {table1};
            var result = maker.Build(path, template);
            var newpath = "2.docx";
            File.WriteAllBytes(newpath,result);
        }
    }
}
