using System;

namespace WordExporter.Core.WordManipulation.Support
{
    public class RunMatch
    {
        public RunMatch(DocumentFormat.OpenXml.Wordprocessing.Run run)
        {
            Run = run;
        }

        public DocumentFormat.OpenXml.Wordprocessing.Run Run { get; set; }

        public Int32 RunStartPosition { get; set; }

        public Int32 RunEndPosition { get; set; }
    }
}
