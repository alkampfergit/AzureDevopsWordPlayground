using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using WordExporter.Core.WordManipulation.Support;

namespace WordExporter.Core.WordManipulation
{
    public class WordManipulator : IDisposable
    {
        public WordManipulator(String fileName, Boolean createNew)
        {
            if (!createNew && !File.Exists(fileName))
                throw new ArgumentException($"File {fileName} does not exists and CreateNew is false.");

            if (createNew == false)
                throw new NotSupportedException("Still not able to open word for manipulation");

            _document = WordprocessingDocument.Create(fileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            _mainDocumentPart = _document.AddMainDocumentPart();
            _body = new Body();
            _mainDocumentPart.Document = new Document(_body);

            InitializeStyles();
        }

        private readonly WordprocessingDocument _document;
        private readonly MainDocumentPart _mainDocumentPart;
        private readonly Body _body;

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

        public void InsertWorkItem(WorkItem workItem, Boolean insertPageBreak = true)
        {
            Log.Debug("Adding to word work item [{Id}/{Type}]: {Title}", workItem.Id, workItem.Type.Name, workItem.Title);
            AppendTextWithStyle("workItemTitle", $"{workItem.Id}: {workItem.Title}");
            _body.Append(
                new Paragraph(
                new Run(
                    new Break() { Type = BreakValues.Page })));
        }

        #endregion

        #region Adding Text helpers

        private void AppendTextWithStyle(String styleId, String text)
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
                    stylesPart = docPart.StylesWithEffectsPart;
                else
                    stylesPart = docPart.StyleDefinitionsPart;

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
            StyleProperties properties )
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
