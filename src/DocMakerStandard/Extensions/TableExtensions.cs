using DocumentFormat.OpenXml.Wordprocessing;

namespace DocMaker.Extensions
{
    public static class TableExtensions
    {
        public static TableRow CloneRow(this TableRow originalRow)
        {
            TableRow newRow = (TableRow)originalRow.CloneNode(true);

            TableRowProperties trh = newRow.OfType<TableRowProperties>().FirstOrDefault();
            if (trh != null)
            {
                TableRowHeight height = trh.OfType<TableRowHeight>().FirstOrDefault();
                if (height == null)
                {
                    TableRowHeight newHeight = new TableRowHeight();
                    newHeight.Val = 0;
                    var curProp = trh.OfType<TableProperties>().FirstOrDefault();
                    if (curProp != null)
                    {
                        curProp.Append(newHeight);
                    }
                    else
                    {
                        TableRowProperties newProp = new TableRowProperties();
                        newProp.Append(newHeight);
                        newRow.Append(newProp);
                    }
                }
                else
                {
                    height.Val = 0;
                }
            }
            else
            {
                TableRowProperties newProp = new TableRowProperties();
                TableRowHeight newHeight = new TableRowHeight();
                newHeight.Val = 0;
                newProp.Append(newHeight);
                newRow.Append(newProp);
            }

            return newRow;
        }
    }
}