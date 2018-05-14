using System.Collections.Generic;

namespace DocMaker.Domain
{
    public class DocTable
    {
        public DocTable()
        {
            Rows = new List<DocTableRow>();
        }

        public IList<DocTableRow> Rows { get; set; }

        public void AddRow(string[] cells)
        {
            var row = new DocTableRow { Cells = cells };
            Rows.Add(row);
        }
    }
}