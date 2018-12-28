using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace WordExporter.Core.WordManipulation
{
    public static class WordManipulationExtensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
                xdoc = XDocument.Load(xr);
            part.AddAnnotation(xdoc);
            return xdoc;
        }

        public static void AppendOtherWordFile(this Document document, String wordFilePath, Boolean addPageBreak = true)
        {
            if (addPageBreak)
            {
                document.AddPageBreak();
            }
            MainDocumentPart mainPart = document.MainDocumentPart;
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
        }

        public static void AddPageBreak(this Document document)
        {
            Body body = document.MainDocumentPart.Document.Body;
            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Break() { Type = BreakValues.Page });
        }
    }
}
