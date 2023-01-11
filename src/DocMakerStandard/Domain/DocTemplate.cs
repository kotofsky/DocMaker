namespace DocMaker.Domain
{
    public class DocTemplate
    {
        public DocTable[] Tables { get; set; }

        public IDictionary<string,string> FieldsCollection { get; set; }
    }
}