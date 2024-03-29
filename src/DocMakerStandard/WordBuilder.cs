﻿using DocMaker.Domain;
using DocMaker.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocMaker
{
    public class WordBuilder
    {
        public DocTemplate CreateTemplate()
        {
            return new DocTemplate();
        }

        public Stream Build(Stream stream, DocTemplate template)
        {
            return Generate(stream, template.FieldsCollection, template.Tables);
        }

        public byte[] Build(string templatePath, DocTemplate template)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"File not found at this path: {templatePath}");

            return Generate(templatePath, template.FieldsCollection, template.Tables);
        }


        private Stream Generate(Stream stream, IDictionary<string, string> fields, DocTable[] docTables = null)
        {
            var resultStream = new MemoryStream();
            stream.CopyTo(resultStream);
            resultStream.Seek(0, SeekOrigin.Begin);
            
            using (var doc = WordprocessingDocument.Open(resultStream, true))
            {
                foreach (var key in fields.Keys)
                {
                    FillContentControls(doc, key, fields[key]);
                }

                if (docTables != null)
                    FillTables(doc, docTables);
            }

            return resultStream;
        }


        private byte[] Generate(string templatePath, IDictionary<string, string> fields, DocTable[] docTables = null)
        {
            byte[] templateData = File.ReadAllBytes(templatePath);
            
            var templateStream = new MemoryStream();
            templateStream.Write(templateData, 0, templateData.Length);
            templateStream.Seek(0, SeekOrigin.Begin);
            
            using (var doc = WordprocessingDocument.Open(templateStream, true))
            {
                foreach (var key in fields.Keys)
                {
                    FillContentControls(doc, key, fields[key]);
                }

                if (docTables != null)
                    FillTables(doc, docTables);
            }

            return templateStream.ToArray();
        }

        private void FillTables(WordprocessingDocument document, DocTable[] tables)
        {
            // get WORD tables
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
                            GridSpan gs = new GridSpan();
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
        private void FillContentControls(WordprocessingDocument document, string name, string value)
        {
            List<SdtElement> blocks = new List<SdtElement>();
            blocks.AddRange(document.MainDocumentPart.Document.Body.Descendants<SdtElement>());
            IEnumerable<SdtElement> matchBlocks = blocks.Where(x =>
            {
                var alias = x.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                return alias != null && alias.Val.Value == name;
            });


            foreach (SdtElement block in matchBlocks)
            {
                if (block is SdtRun)
                {
                    SdtContentRun contRun = block.Descendants<SdtContentRun>().First();

                    Run run = contRun.Descendants().OfType<Run>().FirstOrDefault();
                    if (run == null)
                    {
                        run = new Run();

                        if (!run.Elements<RunProperties>().Any())
                        {
                            run.PrependChild(new RunProperties());
                        }

                        RunProperties runProp = run.RunProperties;
                        if (runProp.RunStyle == null)
                            runProp.RunStyle = new RunStyle();

                        runProp.RunStyle.Val = (block.SdtProperties.FirstChild.FirstChild as RunStyle)?.Val;
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
                    Paragraph contentFieldParagraph = block.Descendants().OfType<Paragraph>().FirstOrDefault();
                    Run run = contentFieldParagraph.OfType<Run>().FirstOrDefault();
                    if (run == null)
                    {
                        run = new Run();
                        SetContentControlValue(run, value);
                        contentFieldParagraph.Append(run);
                    }

                    SetContentControlValue(run, value);

                }
            }
        }

        private void SetContentControlValue(Run run, string value)
        {
            if (string.IsNullOrEmpty(value))
                return;
            IEnumerable<string> lines = value.Contains('\n') ? value.Split("\n".ToCharArray()) : new string[] { value };
            for (int i = 0; i < lines.Count(); i++)
            {
                string line = lines.ElementAt(i);
                line = line.Trim('\n', '\r');
                Text text = new Text(lines.ElementAt(i));
                run.Append(text);
                if (i < lines.Count() - 1)
                {
                    run.Append(new Break());
                }
            }
        }
    }
}