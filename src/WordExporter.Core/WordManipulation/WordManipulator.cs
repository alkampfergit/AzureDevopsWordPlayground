using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using WordExporter.Core.WordManipulation.Support;
using WordExporter.Core.WorkItems;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WordExporter.Core.WordManipulation
{
    public class WordManipulator : IDisposable
    {
        public WordManipulator(String fileName, Boolean createNew)
        {
            if (!createNew && !File.Exists(fileName))
            {
                throw new ArgumentException($"File {fileName} does not exists and CreateNew is false.");
            }

            if (createNew)
            {
                _document = WordprocessingDocument.Create(fileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                _mainDocumentPart = _document.AddMainDocumentPart();
                _body = new Body();
                _mainDocumentPart.Document = new Document(_body);
                InitializeStyles();
            }
            else
            {
                _document = WordprocessingDocument.Open(fileName, true);
                _mainDocumentPart = _document.MainDocumentPart;
                _body = _mainDocumentPart.Document.Body;
            }
        }

        private readonly WordprocessingDocument _document;
        private readonly MainDocumentPart _mainDocumentPart;
        private readonly Body _body;

        public WordprocessingDocument Document => _document;
        public Body DocumentBody => _body;

        #region IDisposable Support

        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _document?.Dispose();
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion

        #region Initialization

        private void InitializeStyles()
        {
            var styleDefinitionPart = AddStylesPartToPackage();
            CreateAndAddParagraphStyle(styleDefinitionPart, "workItemTitle", "Work Item Title", new StyleProperties()
            {
                Bold = true,
                FontSize = 24
            });

            CreateAndAddParagraphStyle(styleDefinitionPart, "workItemBody", "Work Item Title", new StyleProperties()
            {
                Bold = true,
                FontSize = 12
            });
        }

        public StyleDefinitionsPart AddStylesPartToPackage()
        {
            StyleDefinitionsPart part;
            part = _document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        #endregion  

        #region Manipulation of the word document to create data

        /// <summary>
        /// Insert a simple work item, it will use all the fields from work item
        /// and perform substitution, then append the new work item to the current
        /// document.
        /// </summary>
        /// <param name="workItem"></param>
        /// <param name="workItemTemplateFile"></param>
        /// <param name="insertPageBreak"></param>
        /// <param name="startingParameters">These parameters will be added to dictionary
        /// with all fields of work item.</param>
        public void InsertWorkItem(
            WorkItem workItem,
            String workItemTemplateFile,
            Boolean insertPageBreak = true,
            Dictionary<string, object> startingParameters = null)
        {
            //ok we need to open the template, give it a new name, perform substitution and finally append to the existing document
            var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            File.Copy(workItemTemplateFile, tempFile, true);
            startingParameters = startingParameters ?? new Dictionary<string, object>();
            using (WordManipulator m = new WordManipulator(tempFile, false))
            {
                Dictionary<string, object> tokenList = workItem.CreateDictionaryFromWorkItem();
                if (startingParameters != null)
                {
                    foreach (var parameter in startingParameters)
                    {
                        tokenList[parameter.Key] = parameter.Value;
                    }
                }
                m.SubstituteTokens(tokenList);
            }

            AppendOtherWordFile(tempFile, insertPageBreak);
            File.Delete(tempFile);
        }

        #endregion

        #region Template handling

        public WordManipulator AppendOtherWordFile(String wordFilePath, Boolean addPageBreak = true)
        {
            MainDocumentPart mainPart = _document.MainDocumentPart;
            string altChunkId = "AltChunkId" + Guid.NewGuid().ToString();
            AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using (FileStream fileStream = File.Open(wordFilePath, FileMode.Open))
            {
                chunk.FeedData(fileStream);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;
                mainPart.Document
                    .Body
                    .InsertAfter(altChunk, mainPart.Document.Body
                    .Elements().LastOrDefault());
                mainPart.Document.Save();
            }
            if (addPageBreak)
            {
                _body.Append(
                   new Paragraph(
                   new Run(
                       new Break() { Type = BreakValues.Page })));
            }
            return this;
        }

        /// <summary>
        /// This does the very same stuff as <see cref="ReplaceInDocument(Dictionary{string, string}, bool)"/>
        /// but it uses different internal technique to replace the document.
        /// </summary>
        /// <param name="tokenList"></param>
        /// <param name="blankIfMissing"></param>
        /// <returns></returns>
        public WordManipulator SubstituteTokens(
            Dictionary<string, Object> tokenList)
        {
            var realReplaceList = tokenList.ToDictionary(_ => CreateSubstitutionTokenFromName(_.Key), _ => _.Value);

            var body = _document.MainDocumentPart.Document.Body;
            SubstituteInParagraph(realReplaceList, body.Descendants<Paragraph>());

            foreach (var header in _document.MainDocumentPart.HeaderParts)
            {
                SubstituteInParagraph(realReplaceList, header.RootElement.Descendants<Paragraph>());
            }

            foreach (var footer in _document.MainDocumentPart.FooterParts)
            {
                SubstituteInParagraph(realReplaceList, footer.RootElement.Descendants<Paragraph>());
            }

            return this;
        }

        /// <summary>
        /// perform substitution in document content
        /// </summary>
        /// <param name="realReplaceList">This is a list that contains the tag enclosed in {{}}
        /// as key, and can contain string or stream as values.</param>
        /// <param name="paragraphs"></param>
        private void SubstituteInParagraph(
            Dictionary<string, Object> realReplaceList,
            IEnumerable<Paragraph> paragraphs)
        {
            foreach (var paragraph in paragraphs)
            {
                //replace runs with entire code in it, it could happen that a single run
                //contains the entire text.
                var entireRuns = paragraph.Descendants<Run>().ToList();
                foreach (var run in entireRuns)
                {
                    String innerText = run.InnerText;

                    foreach (var replace in realReplaceList)
                    {
                        //Replace text only if the replace is a string.
                        if (replace.Value is String)
                        {
                            //perform a real replace 
                            if (innerText.Contains(replace.Key))
                            {
                                innerText = innerText.Replace(replace.Key, replace.Value as String);
                            }
                        }
                    }

                    //something is changed?
                    if (run.InnerText != innerText)
                    {
                        var newRun = new Run(new Text(innerText));
                        CopyPropertiesFromRun(run, newRun);
                        paragraph.ReplaceChild(newRun, run);
                    }
                }

                //Lets look for each key if it is found in the paragraph and it needs to be replaced.
                //This is the real code that also replace images passed as a stream in dictionary
                foreach (var replace in realReplaceList)
                {
                    //each time we restart we could have changed runs so we re-execute the linq query.
                    Match match;
                    do
                    {
                        //Each cycle runs can be changed
                        List<RunMatch> runs = GetAllRunMatches(paragraph);
                        var paragraphInnerText = paragraph.InnerText;
                        match = Regex.Match(paragraphInnerText, @"\{\{" + replace.Key.Trim('}', '{') + @"(\:[0-9a-zA-Z_-]*?)?\}\}", RegexOptions.IgnoreCase);
                        if (match?.Success == true)
                        {
                            //ok we found a match, we need to grab all the run that encompass this match
                            var start = match.Index;
                            var end = start + match.Length;

                            var runThatMatches = runs.Where(_ =>
                                (_.RunStartPosition >= start && _.RunStartPosition < end)
                                || (_.RunEndPosition > start && _.RunEndPosition <= end)
                            ).ToList();

                            //we have three part, whatever is before the  {{ the tag and then whatever is after }}
                            //in the end we have three runs
                            Run runBefore;
                            var textOfFirstRun = runThatMatches.First().Run.InnerText;
                            var startTokenPosition = textOfFirstRun.IndexOf("{{");
                            if (startTokenPosition > 0)
                            {
                                runBefore = new Run(new Text(textOfFirstRun.Substring(0, startTokenPosition)));
                            }
                            else
                            {
                                runBefore = new Run(new Text(String.Empty));
                            }

                            //now we will add the real replace

                            //finally we can append the trailing space
                            Run runAfter;

                            var textOfLastRun = runThatMatches.Last().Run.InnerText;
                            var endTokenPosition = textOfLastRun.IndexOf("}}");
                            if (endTokenPosition <= textOfLastRun.Length - 3)
                            {
                                runAfter = new Run(new Text(textOfLastRun.Substring(endTokenPosition + 2)));
                            }
                            else
                            {
                                runAfter = new Run(new Text(String.Empty));
                            }

                            //now the content, version 1, only text
                            Object value = replace.Value;
                            OpenXmlElement contextRun = CreateElementFromValue(value, match.Value);

                            var firstRunToReplace = runThatMatches.First().Run;
                            CopyPropertiesFromRun(firstRunToReplace, runBefore);
                            CopyPropertiesFromRun(firstRunToReplace, contextRun as Run);
                            CopyPropertiesFromRun(firstRunToReplace, runAfter);

                            if (contextRun is AltChunk)
                            {
                                paragraph.Parent.ReplaceChild(contextRun, paragraph);
                                //_body.RemoveChild(paragraph);
                                break; //no more replace in this paragraph, html will replace everything
                            }
                            else
                            {
                                paragraph.InsertBefore(runBefore, firstRunToReplace);
                                paragraph.InsertBefore(contextRun, firstRunToReplace);
                                paragraph.InsertBefore(runAfter, firstRunToReplace);

                                foreach (var runToRemove in runThatMatches)
                                {
                                    paragraph.RemoveChild(runToRemove.Run);
                                }
                            }
                        }
                    } while (match?.Success == true);
                }
            }
        }

        private OpenXmlElement CreateElementFromValue(object value, String match)
        {
            if (value == null)
            {
                return new Run(new Text(String.Empty));
            }

            switch (value)
            {
                case String str:
                    return new Run(new Text(str));

                case ImageSubstitution imageSubstitution:
                    return CreateRunFromImage(imageSubstitution, match);

                case HtmlSubstitution htmlSubstitution:
                    return CreateChunkForHtmlPage(htmlSubstitution.HtmlValue);

                default:
                    throw new NotSupportedException($"Element of type {value.GetType().FullName} is not valid for substitution");
            }
        }

        private Run CreateRunFromImage(ImageSubstitution imageSubstitution, String match)
        {
            var imageSplitted = match.Trim('{', '}').Split(':');
            if (imageSplitted.Length == 2)
            {
                if (Int32.TryParse(imageSplitted[1], out Int32 width))
                {
                    imageSubstitution.ResizeToWidth(width);
                }
            }

            ImagePart imagePart;
            MainDocumentPart mainPart = _document.MainDocumentPart;
            using (var ms = imageSubstitution.GetImageStream())
            {
                imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                imagePart.FeedData(ms);
            }
            Int64 cx = 9525 * imageSubstitution.Image.Width;
            Int64 cy = 9525 * imageSubstitution.Image.Height;

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = cx, Cy = cy },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = 1U,
                             Name = Guid.NewGuid().ToString(),
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = Guid.NewGuid().ToString(),
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = mainPart.GetIdOfPart(imagePart),
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = cx, Cy = cy }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle })
                                         )
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                         EditId = "50D07946"
                     });

            return new Run(element);
        }

        private static void CopyPropertiesFromRun(Run originalRun, Run run)
        {
            if (originalRun != null && run != null)
            {
                RunProperties runProperties = originalRun.Descendants<RunProperties>().FirstOrDefault();
                if (runProperties != null && run != null)
                {
                    var copy = runProperties.CloneNode(true);
                    run.InsertAt(copy, 0);
                }
            }
        }

        private static List<RunMatch> GetAllRunMatches(Paragraph paragraph)
        {
            var runs = paragraph.Descendants<Run>()
                 .Select(_ => new RunMatch(_))
                 .ToList();
            Int32 starts = 0;
            foreach (var element in runs)
            {
                element.RunStartPosition = starts;
                starts += element.Run.InnerText.Length;
                element.RunEndPosition = starts;
            }

            return runs;
        }

        private static string CreateSubstitutionTokenFromName(String tokenName)
        {
            return "{{" + tokenName + "}}";
        }

        #endregion

        #region Adding Text helpers

        private WordManipulator AppendTextWithStyle(String styleId, String text)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties properties = new ParagraphProperties();
            properties.ParagraphStyleId = new ParagraphStyleId()
            {
                Val = styleId,
            };
            p.AppendChild(properties);
            p.AppendChild(new Run(new Text(text)));
            _body.Append(p);
            return this;
        }

        public WordManipulator AppendHtml(OpenXmlElement refChild, String htmlPage)
        {
            AltChunk altChunk = CreateChunkForHtmlPage(htmlPage);
            _document.MainDocumentPart.Document.Body.InsertAfter(altChunk, refChild);
            return this;
        }

        private AltChunk CreateChunkForHtmlPage(string htmlPage)
        {
            var realHtml = $"<html><head></head><body>{htmlPage}</body></html>";
            string altChunkId = "myid" + Guid.NewGuid().ToString();
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(realHtml)))
            {
                // Create alternative format import part.
                AlternativeFormatImportPart formatImportPart = _document.MainDocumentPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.Html,
                    altChunkId);

                // Feed HTML data into format import part (chunk).
                formatImportPart.FeedData(ms);
            }
            var altChunk = new AltChunk();
            altChunk.Id = altChunkId;
            return altChunk;
        }

        //AppendHtml(p1, $"<html><head></head><body>{workItem.Description}</body></html>");
        //public WordManipulator AppendHtml(OpenXmlElement refChild, String htmlPage)
        //{
        //    string altChunkId = "myid" + Guid.NewGuid().ToString();
        //    MainDocumentPart mainDocPart = _document.MainDocumentPart;

        //    using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlPage)))
        //    {
        //        // Create alternative format import part.
        //        AlternativeFormatImportPart formatImportPart = mainDocPart.AddAlternativeFormatImportPart(
        //            AlternativeFormatImportPartType.Html,
        //            altChunkId);

        //        // Feed HTML data into format import part (chunk).
        //        formatImportPart.FeedData(ms);
        //        AltChunk altChunk = new AltChunk();
        //        altChunk.Id = altChunkId;

        //        mainDocPart.Document.Body.InsertAfter(altChunk, refChild);
        //    }
        //    return this;
        //}

        #endregion

        #region Table

        /// <summary>
        /// <para>
        /// This is a really rudimental method to fill a table with content, it assumes
        /// that the opened document has only one table, and it will fill the table.
        /// with data.
        /// </para>
        /// <para>
        /// If the table already have rows, all rows will be deleted, but the first row
        /// will be used to maintain the formatting styles. This means that if you want
        /// formatting to be maintained you should simply insert a row with or without data
        /// formatted as you like.
        /// </para>
        /// </summary>
        /// <param name="skipHeader">If true it will skip the first line of the table.</param>
        /// <param name="data">this is a matrix, expressed by a couple of IEnumerable to simplify
        /// data passing</param>
        /// <returns></returns>
        public WordManipulator FillTable(
            Boolean skipHeader,
            IEnumerable<IEnumerable<Object>> data)
        {
            var table = _document.MainDocumentPart.Document.Body
                .Descendants<Table>()
                .FirstOrDefault();
            if (table != null)
            {
                //remove every rows but first save the first two rows for the formatting.
                var rows = table.Elements<TableRow>().ToList();
                Int32 skip = skipHeader ? 1 : 0;
                foreach (var row in rows.Skip(skip))
                {
                    row.Remove();
                }

                Boolean isFirstRow = true;
                foreach (var dataRow in data)
                {
                    TableRow row = null;
                    TableRow templateRow;
                    //template rows depends from skipping or not skipping the header template.
                    if (skipHeader)
                    {
                        //header is skipped, take the second row if present.
                        templateRow = rows.Skip(1).FirstOrDefault();
                    }
                    else
                    {
                        //we do not want to skip header
                        var rowToSkip = isFirstRow || rows.Count < 2 ? 0 : 1;
                        templateRow = rows.Skip(rowToSkip).FirstOrDefault();
                    }
                    if (templateRow == null)
                    {
                        row = new TableRow();
                        foreach (var dataCell in dataRow)
                        {
                            var cell = new TableCell();

                            // Specify the table cell content.
                            cell.Append(new Paragraph(new Run(new Text(dataCell.ToString()))));

                            // Append the table cell to the table row.
                            row.Append(cell);
                        }
                    }
                    else
                    {
                        row = (TableRow)templateRow.CloneNode(true);
                        //Grab all the run style of first row to copy on all subsequence cell.
                        var runs = templateRow.Descendants<TableCell>()
                            .Select(_ => _.Descendants<Run>().FirstOrDefault())
                            .ToList();
                        var cells = row.Descendants<TableCell>().ToList();
                        Int32 cellIndex = 0;
                        foreach (var dataCell in dataRow)
                        {
                            if (cellIndex < cells.Count)
                            {
                                var cell = cells[cellIndex];
                                var run = runs[cellIndex];

                                // Specify the table cell content.
                                Run runToAdd = new Run(new Text(dataCell.ToString()));
                                CopyPropertiesFromRun(run, runToAdd);

                                //we can  have two distinct situation, we have or we do not have paragraph
                                var paragraph = cell.Descendants<Paragraph>().FirstOrDefault();
                                if (paragraph == null)
                                {
                                    cell.Append(new Paragraph(runToAdd));
                                }
                                else
                                {
                                    paragraph.RemoveAllChildren<Run>();
                                    paragraph.Append(runToAdd);
                                }
                            }
                            cellIndex++;
                        }
                    }
                    table.Append(row);
                    isFirstRow = false;
                }
            }
            return this;
        }

        /// <summary>
        /// A composite table is a table where we have usually an header, then a first row that have
        /// cells populated with substitution value. We need to replicate the row with the substitution
        /// for each cell value.
        /// </summary>
        /// <param name="skipHeader"></param>
        /// <param name="data">A series of dictionaries, each dictionary contains
        /// substitution values for an entire row.</param>
        /// <returns></returns>
        public WordManipulator FillCompositeTable(
          Boolean skipHeader,
          IEnumerable<Dictionary<String, Object>> data)
        {
            var table = _document.MainDocumentPart.Document.Body
                .Descendants<Table>()
                .FirstOrDefault();
            if (table != null)
            {
                //remove every rows but keep formatting row.
                var rows = table.Elements<TableRow>().ToList();
                Int32 skip = skipHeader ? 1 : 0;

                TableRow templateRow = rows.Skip(skip).FirstOrDefault();
                table.RemoveChild(templateRow);

                if (templateRow != null)
                {
                    foreach (var dataRow in data)
                    {
                        var row = (TableRow)templateRow.CloneNode(true);

                        //Grab all the run style of first row to copy on all subsequence cell.
                        var paragraph = row.Descendants<Paragraph>();
                        var realReplaceList = dataRow.ToDictionary(_ => CreateSubstitutionTokenFromName(_.Key), _ => _.Value);

                        SubstituteInParagraph(realReplaceList, paragraph);
                        table.Append(row);
                    }
                }
            }
            return this;
        }

        /// <summary>
        /// We have a table with an header (optional) and the first row, where in each cell we 
        /// specify the field of the work item we want to include in the cell.
        /// </summary>
        /// <param name="skipHeader"></param>
        /// <param name="workItems"></param>
        /// <returns></returns>
        public WordManipulator FillTableWithSingleFieldWorkItems(
            Boolean skipHeader,
            IEnumerable<WorkItem> workItems)
        {
            var table = _document.MainDocumentPart.Document.Body
                .Descendants<Table>()
                .FirstOrDefault();
            if (table != null)
            {
                //remove every rows but first save the first two rows for the formatting.
                var rows = table.Elements<TableRow>().ToList();
                Int32 skip = skipHeader ? 1 : 0;
                var rowWithField = rows.Skip(skip).First();
                List<List<String>> workItemCellsData = new List<List<string>>();
                List<String> cellFields = new List<string>();
                var cells = rowWithField.Descendants<TableCell>().ToList();
                foreach (var dataCell in cells)
                {
                    //ok for each cell we need to grab the field name
                    //in this first version we support only a field for each column
                    var dataCellText = dataCell.InnerText;
                    var match = Regex.Match(dataCellText, @"\{\{(?<name>.*?)\}\}");
                    if (match.Success)
                    {
                        cellFields.Add(match.Groups["name"].Value);
                    }
                    else
                    {
                        cellFields.Add("");
                    }
                }
                foreach (var workItem in workItems)
                {
                    List<String> row = new List<string>();
                    var properties = workItem.CreateDictionaryFromWorkItem();
                    foreach (var field in cellFields)
                    {
                        if (properties.TryGetValue(field, out var value))
                        {
                            row.Add(value as String);
                        }
                        else
                        {
                            row.Add(String.Empty);
                        }
                    }
                    workItemCellsData.Add(row);
                }

                FillTable(true, workItemCellsData);
            }
            return this;
        }

        /// <summary>
        /// This is similar to <see cref="FillTableWithSingleFieldWorkItems(bool, IEnumerable{WorkItem})"/>
        /// but with this version the routine will perform a complete substitution in each cell
        /// so you can have multiple value in cells. This is more time consuming that
        /// <see cref="FillTableWithSingleFieldWorkItems(bool, IEnumerable{WorkItem})"/>
        /// </summary>
        /// <param name="skipHeader"></param>
        /// <param name="workItems"></param>
        /// <returns></returns>
        public WordManipulator FillTableWithCompositeWorkItems(
            Boolean skipHeader,
            IEnumerable<WorkItem> workItems)
        {
            if (!workItems.Any())
                return this;

            List<Dictionary<String, Object>> workItemCellsData = new List<Dictionary<String, Object>>();
            List<Int32> parentList = workItems
                .Select(w => GetParentLink(w))
                .Where(l => l != null)
                .Select(l => l.RelatedWorkItemId)
                .ToList();

            Dictionary<int, WorkItem> parentWorkItems = GetParentsInformation(parentList);

            List<Int32> granParentList = parentWorkItems.Values
               .Select(w => GetParentLink(w))
               .Where(l => l != null)
               .Select(l => l.RelatedWorkItemId)
               .ToList();

            Dictionary<int, WorkItem> granParentWorkItems = GetParentsInformation(granParentList);

            foreach (var workItem in workItems)
            {
                var parameters = workItem.CreateDictionaryFromWorkItem();
                RelatedLink parentLink = GetParentLink(workItem);
                parameters["parent.title"] = String.Empty;
                parameters["parent.id"] = String.Empty;
                parameters["parent.parent.title"] = String.Empty;
                parameters["parent.parent.id"] = String.Empty;
                if (parentWorkItems.TryGetValue(parentLink?.RelatedWorkItemId ?? 0, out var parent))
                {
                    parameters["parent.title"] = parent.Title;
                    parameters["parent.id"] = parent.Id;
                    var granParentLink = GetParentLink(parent);
                    if (granParentWorkItems.TryGetValue(granParentLink?.RelatedWorkItemId ?? 0, out var granParent))
                    {
                        parameters["parent.parent.title"] = granParent.Title;
                        parameters["parent.parent.id"] = granParent.Id;
                    }
                }
                workItemCellsData.Add(parameters);
            }
            return FillCompositeTable(skipHeader, workItemCellsData);
        }

        private static Dictionary<int, WorkItem> GetParentsInformation(List<int> parentList)
        {
            //optimize, load all parents in a dictionary with a single query
            Dictionary<Int32, WorkItem> parentWorkItems = new Dictionary<int, WorkItem>();
            if (parentList.Count > 0)
            {
                //ok now we need to grab all parent link, just to grab 
                var query = $@"SELECT
    [System.Id],
    [System.Title]
FROM workitems
WHERE [System.Id] IN ({String.Join(",", parentList)})
ORDER BY [System.Id]
";
                //ok, query all the parents
                parentWorkItems = ConnectionManager.Instance.WorkItemStore.Query(query)
                    .OfType<WorkItem>()
                    .ToDictionary(w => w.Id);
            }

            return parentWorkItems;
        }

        private static RelatedLink GetParentLink(WorkItem workItem)
        {
            return workItem
                .Links
                .OfType<RelatedLink>()
                .SingleOrDefault(l => l.LinkTypeEnd.Name == "Parent");
        }

        #endregion

        #region Helpers

        private Dictionary<String, String> _styleCache = new Dictionary<string, string>();

        public static XDocument ExtractStylesPart(
            string fileName,
            bool getStylesWithEffectsPart = true)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            // Open the document for read access and get a reference.
            using (var document =
                WordprocessingDocument.Open(fileName, false))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the
                // stylesPart variable.
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                {
                    stylesPart = docPart.StylesWithEffectsPart;
                }
                else
                {
                    stylesPart = docPart.StyleDefinitionsPart;
                }

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                      stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                    }
                }
            }
            // Return the XDocument instance.
            return styles;
        }

        private String GetStyleIdFromName(String styleName)
        {
            if (!_styleCache.TryGetValue(styleName, out var styleId))
            {
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                var stylePart = _document
                    .MainDocumentPart
                    .StyleDefinitionsPart
                    .Styles;

                var styles = stylePart
                    .Descendants<Style>();

                styleId = styles
                    .First(_ => _.StyleName.Val.Value == styleName)
                    .StyleId;

                _styleCache.Add(styleName, styleId);
            }
            return styleId;
        }

        public void CreateAndAddParagraphStyle(
            StyleDefinitionsPart styleDefinitionsPart,
            string styleid,
            string stylename,
            StyleProperties properties)
        {
            // Access the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;
            if (styles == null)
            {
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();
            }

            // Create a new paragraph style element and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            AutoRedefine autoredefine1 = new AutoRedefine() { Val = OnOffOnlyValues.Off };
            BasedOn basedon1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountChar" };
            Locked locked1 = new Locked() { Val = OnOffOnlyValues.Off };
            PrimaryStyle primarystyle1 = new PrimaryStyle() { Val = OnOffOnlyValues.On };
            StyleHidden stylehidden1 = new StyleHidden() { Val = OnOffOnlyValues.Off };
            SemiHidden semihidden1 = new SemiHidden() { Val = OnOffOnlyValues.Off };
            StyleName styleName1 = new StyleName() { Val = stylename };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            UIPriority uipriority1 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhidewhenused1 = new UnhideWhenUsed() { Val = OnOffOnlyValues.On };

            style.Append(autoredefine1);
            style.Append(basedon1);
            style.Append(linkedStyle1);
            style.Append(locked1);
            style.Append(primarystyle1);
            style.Append(stylehidden1);
            style.Append(semihidden1);
            style.Append(styleName1);
            style.Append(nextParagraphStyle1);
            style.Append(uipriority1);
            style.Append(unhidewhenused1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = properties.FontSize.ToString() };

            if (properties.Bold)
            {
                styleRunProperties.Append(new Bold());
            }
            styleRunProperties.Append(color1);
            styleRunProperties.Append(font1);
            styleRunProperties.Append(fontSize1);
            styleRunProperties.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        #endregion
    }
}
