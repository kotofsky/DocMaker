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
            
            //stream example
            //using (var fileStream = new FileStream(path,FileMode.Open))
            //{
            //    using (var result = maker.Build(fileStream, template))
            //    {
            //        var memStream = new MemoryStream();
            //        result.Seek(0, SeekOrigin.Begin);
            //        result.CopyTo(memStream);

            //        File.WriteAllBytes("2.docx",memStream.ToArray());
            //    }
            //}

            //byte[] example
            var byteResult = maker.Build(path, template);
            File.WriteAllBytes("2.docx", byteResult);



        }
    }
}
