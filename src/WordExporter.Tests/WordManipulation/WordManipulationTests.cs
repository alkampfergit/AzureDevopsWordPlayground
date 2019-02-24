using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.WordManipulation;

namespace WordExporter.Tests.WordManipulation
{
    [TestFixture]
    public class WordManipulationTests
    {
        [TearDown]
        public void TearDown()
        {
            foreach (var file in _generatedFileNames)
            {
                try
                {
                    File.Delete(file);
                }
#pragma warning disable RCS1075 // Avoid empty catch clause that catches System.Exception.
#pragma warning disable S108 // Nested blocks of code should not be left empty
                catch (Exception)
                {
                }
#pragma warning restore S108 // Nested blocks of code should not be left empty
#pragma warning restore RCS1075 // Avoid empty catch clause that catches System.Exception.
            }
        }

        [Test]
        public void TestNumbering()
        {
            var baseFile = CopyTestFileIntoTempDirectory("base.docx");
            var header1 = CopyTestFileIntoTempDirectory("HeaderTitle1.docx");
            var header2 = CopyTestFileIntoTempDirectory("HeaderTitle2.docx");

            using (var wm = new WordManipulator(baseFile, false))
            {
                wm.AppendOtherWordFile(AppendTitle(header1, "Title 1"), false);
                wm.AppendOtherWordFile(AppendTitle(header2, "Subtitle 1.1"), false);
                wm.AppendOtherWordFile(AppendTitle(header2, "Subtitle 1.2"), false);
                wm.AppendOtherWordFile(AppendTitle(header1, "Title 2"), false);
                wm.AppendOtherWordFile(AppendTitle(header2, "Subtitle 2.1"), false);
            }

            //Open(baseFile);
        }

        private void Open(string baseFile)
        {
            if (Environment.UserInteractive)
            {
                var process = System.Diagnostics.Process.Start(baseFile);
                process.WaitForExit();
            }
        }

        private string AppendTitle(string templateFile, String title)
        {
            using (var wm = new WordManipulator(templateFile, false))
            {
                wm.SubstituteTokens(new Dictionary<string, object>()
                {
                    ["Title"] = title
                });
            }
            return templateFile;
        }

        private readonly HashSet<String> _generatedFileNames = new HashSet<string>();

        private String CopyTestFileIntoTempDirectory(String fileName)
        {
            var testfileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "docs\\" + fileName);
            return CopyIntoTempFile(testfileName);
        }

        private string CopyIntoTempFile(string testfileName)
        {
            var randomTestFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            File.Copy(testfileName, randomTestFileName, true);
            _generatedFileNames.Add(randomTestFileName);
            return randomTestFileName;
        }
    }
}
