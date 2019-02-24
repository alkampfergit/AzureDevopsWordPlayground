using NUnit.Framework;
using System;
using System.IO;
using WordExporter.Core.Templates;
using WordExporter.Tests.Data;

namespace WordExporter.Tests.Templates
{
    [TestFixture]
    public class TemplateManagerTests
    {
        [Test]
        public void Verify_throw_if_folder_is_null()
        {
            Assert.Throws<ArgumentNullException>(() => new TemplateManager(null));
        }

        [Test]
        public void Verify_throw_if_folder_does_not_exists()
        {
            Assert.Throws<ArgumentException>(() => new TemplateManager(GetTemplateFolder("this_does_not_exists")));
        }

        [Test]
        public void Verify_correct_number_of_templates_is_scanned()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            Assert.That(sut.TemplateCount, Is.EqualTo(4));
        }

        [Test]
        public void Verify_throw_if_wrong_template_name_grabbed()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            Assert.Throws<ArgumentException>(() => sut.GetWordDefinitionTemplate("not exsists"));
        }

        [Test]
        public void Verify_enumeration_of_templates()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            Assert.That(sut.GetTemplateNames(), Is.EquivalentTo(new[] { "TemplateA", "TemplateB", "TemplateStructure", "TemplateNumbering" }));
        }

        [Test]
        public void Verify_retrieve_of_template()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            var template = sut.GetWordDefinitionTemplate("TemplateA");
            Assert.That(template, Is.Not.Null);
            Assert.That(template.Name, Is.EqualTo("TemplateA"));

            var expected = Path.Combine(GetTemplateFolder("1"), "TemplateA", "Product Backlog Item.docx");
            Assert.That(template.GetTemplateFor("Product BACKLOG Item"), Is.EqualTo(expected));
        }

        [Test]
        public void Verify_auto_parsing_of_structure_file()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            var template = sut.GetWordDefinitionTemplate("TemplateStructure");
            Assert.That(template, Is.Not.Null);
            Assert.That(template.TemplateDefinition, Is.Not.Null);
            Assert.That(template.TemplateDefinition.AllSections.Length, Is.EqualTo(2));
            Assert.That(template.TemplateDefinition.Parameters.ParameterNames, Is.EquivalentTo(new[] { "test", "blah" }));
        }

        [Test]
        public void Verify_template_with_syntax()
        {
            var sut = new TemplateManager(GetTemplateFolder("1"));
            var template = sut.GetWordDefinitionTemplate("TemplateA");
            Assert.That(template, Is.Not.Null);
            Assert.That(template.Name, Is.EqualTo("TemplateA"));

            var expected = Path.Combine(GetTemplateFolder("1"), "TemplateA", "Product Backlog Item.docx");
            Assert.That(template.GetTemplateFor("Product BACKLOG Item"), Is.EqualTo(expected));
        }

        private static string GetTemplateFolder(String testFolderName)
        {
            return DataManager.GetTemplateFolder(testFolderName);
        }
    }
}
