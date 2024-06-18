namespace DocMaker.Domain;

public class ExcelRow
{
    public int RowIndex { get; set; }

    public Dictionary<string, string> Cells { get; set; } = [];
}
