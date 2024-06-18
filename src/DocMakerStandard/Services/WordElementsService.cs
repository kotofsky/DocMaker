using DocMaker.Domain;
using DocMaker.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocMaker.Services;

internal class WordElementsService
{
    /// <summary>
    /// Creating word tables
    /// </summary>
    /// <param name="document">The whole document</param>
    /// <param name="tables">Inner type with table data</param>
    internal void ProcessWordTables(WordprocessingDocument document, DocTable[] tables)
    {
        // get tables
        Table[] originalTables = document.MainDocumentPart.Document.Body.Elements<Table>().ToArray();

        for (int i = 0; i < tables.Length; i++)
        {
            var clientTable = tables[i];

            Table table = originalTables[i];

            TableRow firstRow = table.Elements<TableRow>().ElementAt(0);

            // iterate each row
            foreach (var tr in clientTable.Rows)
            {
                TableRow newRow = firstRow.CloneRow();
                var cells = tr.Cells;

                int j = 0;
                foreach (var td in cells)
                {
                    if (td == "-")
                    {
                        TableCell firstNewCell = newRow.Elements<TableCell>().First();
                        TableCell toDelCell = newRow.Elements<TableCell>().ToList()[1];
                        toDelCell.Remove();
                        GridSpan gs = new();
                        int countToSkip = cells.Count(x => x == "-") + 1;
                        gs.Val = countToSkip;
                        firstNewCell.TableCellProperties.Append(gs);
                    }
                    else
                    {
                        TableCell newCell = newRow.Elements<TableCell>().ElementAt(j);

                        Text textCop = newCell.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First();

                        // clear all paragraphs 
                        foreach (var paragraph in newCell.Elements<Paragraph>())
                        {
                            foreach (var currun in paragraph.Elements<Run>())
                            {
                                foreach (var curText in currun.Elements<Text>())
                                    curText.Remove();
                            }
                        }

                        newCell.Elements<Paragraph>().First().Elements<Run>().First().AppendChild(textCop);

                        Paragraph par = newCell.Elements<Paragraph>().First();
                        Run run = par.Elements<Run>().First();
                        Text text = run.Elements<Text>().First();
                        text.Text = td;

                        j++;
                    }
                }
                table.AppendChild(newRow);
            }
        }
    }

    /// <summary>
    /// Method for update content-control element
    /// </summary>
    /// <param name = "document" >Word document</param>
    /// <param name = "name" >Name of the field</param>
    /// <param name = "value" >Value of the field</param>
    internal void SetContentControls(WordprocessingDocument document, string name, string value)
    {
        List<SdtElement> blocks = [.. document?.MainDocumentPart?.Document?.Body?.Descendants<SdtElement>()];

        IEnumerable<SdtElement> matchBlocks = blocks.Where(x =>
        {
            SdtAlias? alias = x.SdtProperties?.OfType<SdtAlias>().FirstOrDefault(s => s.Val == name);
            return alias?.Val?.Value == name;
        });

        foreach (SdtElement block in matchBlocks)
        {
            if (block is SdtRun)
            {
                SdtContentRun contRun = block.Descendants<SdtContentRun>().FirstOrDefault() ?? new SdtContentRun();

                var run = contRun?.Descendants()?.OfType<Run>()?.FirstOrDefault();
                if (run == null)
                {
                    run = new Run();

                    if (!run.Elements<RunProperties>().Any())
                    {
                        run.PrependChild(new RunProperties());
                    }

                    RunProperties runProp = run?.RunProperties ?? new RunProperties();
                    runProp.RunStyle ??= new RunStyle();

                    runProp.RunStyle.Val = (block?.SdtProperties?.FirstChild?.FirstChild as RunStyle)?.Val;
                    SetContentControlValue(run, value);
                    contRun.Append(run);
                }
                else
                {
                    SetContentControlValue(run, value);
                }
            }
            else
            {
                Paragraph contentFieldParagraph = block.Descendants().OfType<Paragraph>().FirstOrDefault() ?? new Paragraph();
                Run run = contentFieldParagraph?.OfType<Run>().FirstOrDefault() ?? new Run();
                if (run == null)
                {
                    SetContentControlValue(run, value);
                    contentFieldParagraph.Append(run);
                }

                SetContentControlValue(run, value);
            }
        }
    }

    internal void SetContentControlValue(Run? run, string value)
    {
        if (string.IsNullOrEmpty(value))
            return;

        // Split the value into lines and trim them
        var lines = value.Split(new[] { '\n' }, StringSplitOptions.None)
                         .Select(line => line.Trim('\n', '\r'))
                         .ToList();

        // Iterate through the lines
        for (int i = 0; i < lines.Count; i++)
        {
            run.Append(new Text(lines[i]));

            // Append a break if it's not the last line
            if (i < lines.Count - 1)
            {
                run.Append(new Break());
            }
        }
    }
}
