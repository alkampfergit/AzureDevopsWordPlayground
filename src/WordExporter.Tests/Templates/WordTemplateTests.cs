using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Templates;
using WordExporter.Tests.Data;

namespace WordExporter.Tests.Templates
{
    /// <summary>
    /// Test for <see cref="WordTemplateDefinition"/> class
    /// </summary>
    [TestFixture]
    public class WordTemplateTests
    {
        [Test]
        public void Verify_throws_on_null()
        {
            Assert.Throws<ArgumentNullException>(() => new WordTemplateFolderManager(null));
        }

        [Test]
        public void Verify_scan_folder_grab_name()
        {
            WordTemplateFolderManager sut = CreateSutForTemplate_1_A();
            Assert.That(sut.Name, Is.EqualTo("TemplateA"));
        }

        [Test]
        public void Verify_correctly_grab_docx_name()
        {
            WordTemplateFolderManager sut = CreateSutForTemplate_1_A();
            var expected = Path.Combine(GetTemplate1AFolder(), "Product Backlog Item.docx");
            Assert.That(sut.GetTemplateFor("Product Backlog Item"), Is.EqualTo(expected));
        }

        [Test]
        public void Verify_grab_docx_name_is_not_case_sensitive()
        {
            WordTemplateFolderManager sut = CreateSutForTemplate_1_A();
            var expected = Path.Combine(GetTemplate1AFolder(), "Product Backlog Item.docx");
            Assert.That(sut.GetTemplateFor("Product BACKLOG Item"), Is.EqualTo(expected));
        }

        [Test]
        public void Verify_default_name_if_Work_item_type_does_not_esixts()
        {
            WordTemplateFolderManager sut = CreateSutForTemplate_1_A();
            var expected = Path.Combine(GetTemplate1AFolder(), "WorkItem.docx");
            Assert.That(sut.GetTemplateFor("This type does not exists"), Is.EqualTo(expected));
        }

        private static WordTemplateFolderManager CreateSutForTemplate_1_A()
        {
            var templateA = GetTemplate1AFolder();
            var sut = new WordTemplateFolderManager(templateA);
            return sut;
        }

        private static string GetTemplate1AFolder()
        {
            return Path.Combine(DataManager.GetTemplateFolder("1"), "TemplateA");
        }
    }
}
