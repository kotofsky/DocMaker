using System.Collections.Generic;

namespace DocMaker.Domain
{
    public class DocTemplate
    {
        public DocTable[] Tables { get; set; }

        public Dictionary<string,string> FieldsCollection { get; set; }
    }
}