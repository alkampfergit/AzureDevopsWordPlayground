using NUnit.Framework;
using System.Linq;
using WordExporter.Core.Templates;
using WordExporter.Core.Templates.Parser;

namespace WordExporter.Tests.Templates.Parser
{
    [TestFixture]
    public class BasicParserTests
    {
        [Test]
        public void Basic_Parsing_of_Parameters_returns_parameter_section()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameters]]
    parama
    paramb");
            Assert.That(def.Parameters, Is.Not.Null);
        }

        [Test]
        public void Parsing_two_Sections()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameters]]
    parama
    paramb
[[static]]
    filename: bla.txt");
            Assert.That(def.AllSections.Length, Is.EqualTo(2));
            Assert.That(def.AllSections[0], Is.InstanceOf<ParameterSection>());
            Assert.That(def.AllSections[1], Is.InstanceOf<StaticWordSection>());
        }

        [Test]
        public void Valid_static_section()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[static]]
    filename: bla.docx
    pageBreak: true");
           var staticSection = def.AllSections.Single() as StaticWordSection;
            Assert.That(staticSection.FileName, Is.EqualTo("bla.docx"));
            Assert.That(staticSection.PageBreak, Is.EqualTo(true));
        }

        [Test]
        public void Valid_query_section()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    template/Product Backlog Item: pbix.docx
    template/Bug: bugaa.docx
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.Query, Is.EqualTo("SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'"));
        }

        [Test]
        public void Query_section_parse_specific_workItem_templates()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    template/Product Backlog Item: pbix.docx
    template/Bug: bugaa.docx
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.GetTemplateForWorkItem("Product Backlog Item"), Is.EqualTo("pbix.docx"));
            Assert.That(querySection.GetTemplateForWorkItem("Task"), Is.EqualTo(null));
        }

        [Test]
        public void Basic_Parsing_of_custom_Parameters()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameters]]
    parama
    paramb
");
            Assert.That(def.Parameters.ParameterNames, Is.EquivalentTo(new[] { "parama", "paramb" }));
        }

        [Test]
        public void Valid_query_section_with_table_file()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    tableTemplate: table.docx
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.TableTemplate, Is.EqualTo("table.docx"));
        }
    }
}
