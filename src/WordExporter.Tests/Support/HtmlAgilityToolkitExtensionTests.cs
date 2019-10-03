using HtmlAgilityPack;
using NUnit.Framework;
using WordExporter.Core.Support;

namespace WordExporter.Tests.Support
{
    [TestFixture]
    public class HtmlAgilityToolkitExtensionTests
    {
        [Test]
        public void SelectivelyRemoveSingleTag()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml("<p>paragraph</p><p>multiple</p>");
            doc.RemoveTags(null, "p");

            Assert.That(doc.DocumentNode.InnerHtml, Is.EqualTo("paragraphmultiple"));
        }

        [Test]
        public void SelectivelyRemoveSingleTagWithPrefix()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml("Prefix<p>paragraph</p><p>multiple</p>");
            doc.RemoveTags(null, "p");

            Assert.That(doc.DocumentNode.InnerHtml, Is.EqualTo("Prefixparagraphmultiple"));
        }

        [Test]
        public void SelectivelyRemoveSingleTagWithClosingSubstitution()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml("<p>paragraph</p><p>multiple</p>");
            doc.RemoveTags(doc.CreateElement("br"), "p");

            Assert.That(doc.DocumentNode.InnerHtml, Is.EqualTo("paragraph<br>multiple<br>"));
        }

        [Test]
        public void SelectivelyRemoveSingleTagMaintainInner()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml("<p>paragraph <strong>strong</strong></p><p>multiple</p>");
            doc.RemoveTags(null, "p");

            Assert.That(doc.DocumentNode.InnerHtml, Is.EqualTo("paragraph <strong>strong</strong>multiple"));
        }

        [Test]
        public void RemoveTableLeavingContent()
        {
            var withoutTable = HtmlAgilityToolkitExtension.RemoveTable(htmlWithTable);

            Assert.That(withoutTable.Contains("<span>test2</span>"));
            Assert.That(!withoutTable.Contains("<table>"));
            Assert.That(!withoutTable.Contains("<tr>"));
            Assert.That(!withoutTable.Contains("<td>"));
        }

        private const string htmlWithTable = @"<table><tbody>
<tr><td></td></tr><span><ul><li><span>Atest</span> </li><li><span>test2</span> </li><li><span>Active</span> 
</li><li><span>ASDF</span> </li></tbody></table>";
    }
}
