using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Tests.Data
{
    /// <summary>
    /// Simplify access to test data.
    /// </summary>
    public static class DataManager
    {
        public static String GetTemplateFolder(String templateName)
        {
            return Path.Combine(TestContext.CurrentContext.TestDirectory, "Data", "Templates", templateName);
        }
    }
}
