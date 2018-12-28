using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }

        private WordprocessingDocument _document;
        private readonly MainDocumentPart _mainDocumentPart;

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        #region Manipulation of the word document to create data

        #endregion

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
    }
}
