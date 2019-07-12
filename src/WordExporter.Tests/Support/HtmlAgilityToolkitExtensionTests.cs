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
    }
}
