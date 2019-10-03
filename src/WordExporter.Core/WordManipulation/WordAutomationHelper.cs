using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using Serilog;

namespace WordExporter.Core.WordManipulation
{
    public class WordAutomationHelper : IDisposable
    {
        private readonly String _fileName;
        private readonly Application app;
        private readonly Document doc;

        public WordAutomationHelper(String fileName, Boolean visibility)
        {
            _fileName = fileName;
            app = new Application();
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            app.Visible = visibility;

            Log.Debug("Opening {0} in office", _fileName);
            //it is importantì
            doc = app.Documents.Open(_fileName);
        }

        public void UpdateAllTocs()
        {
            foreach (TableOfContents toc in doc.TablesOfContents)
            {
                toc.Update();
            }

        }

        public String ConvertToPdf()
        {
            try
            {
                Log.Information("About to converting file {0} to pfd", _fileName);
                var destinationPdf = Path.ChangeExtension(_fileName, ".pdf");
                doc.SaveAs2(destinationPdf, WdSaveFormat.wdFormatPDF);
                Log.Debug("File {0} converted to pdf. Closing word", _fileName);
                //doc.Close();
                //app.Quit();
                return destinationPdf;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error converting {0} - {1}", _fileName, ex.Message);
            }
            return null;
        }

        public void Close()
        {
            if (doc != null)
            {
                this.Close(doc);
            }
            if (app != null)
            {
                this.Close(app);
            }
        }

        private void Close(Application app)
        {
            try
            {
                app.Quit();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Unable to close word application {0}", ex.Message);
                //TODO: Try to kill the process.
            }
        }

        private void Close(Document doc)
        {
            try
            {
                doc.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Unable to close document {0}", ex.Message);
            }
        }

        #region IDisposable Support

        private bool disposedValue; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Close();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
