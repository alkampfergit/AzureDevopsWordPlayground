﻿using NUnit.Framework;
using System;
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
            Assert.That(def.ParameterSection, Is.Not.Null);
        }

        [Test]
        public void Basic_Parsing_of_definition_for_excel_Export()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[definition]]
  type=excel
[[parameters]]
  TargetDateStart=2019-06-01
  TargetDateEnd=2019-10-01
[[parameterDefinition]]");
            Assert.That(def.Type, Is.EqualTo(TemplateType.Excel));
        }

        [Test]
        public void Default_template_type_is_word()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameters]]
  TargetDateStart=2019-06-01
  TargetDateEnd=2019-10-01
[[parameterDefinition]]");
            Assert.That(def.Type, Is.EqualTo(TemplateType.Word));
        }

        [Test]
        public void Support_for_base_template()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[definition]]
  type=excel
  baseTemplate=test.xlsx
[[parameters]]
  TargetDateStart=2019-06-01
  TargetDateEnd=2019-10-01
[[parameterDefinition]]");
            Assert.That(def.BaseTemplate, Is.EqualTo ("test.xlsx"));
        }

        [Test]
        public void Basic_Parsing_of_Parameters_with_default_value_returns_parameter_section()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameters]]
    parama=2017-01-01
    paramb=2019-04-31");
            Assert.That(def.ParameterSection, Is.Not.Null);
            var paramA = def.ParameterSection.Parameters["parama"];
            Assert.That(paramA, Is.EqualTo("2017-01-01"));

            var paramB = def.ParameterSection.Parameters["paramb"];
            Assert.That(paramB, Is.EqualTo("2019-04-31"));
        }

        [Test]
        public void Basic_Parsing_of_Parameters_with_allowed_values()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[parameterDefinition]]
    parama=string/A|B|C
    paramb=datetime");
            Assert.That(def.ParameterDefinition, Is.Not.Null);
            var paramA = def.ParameterDefinition["parama"];
            Assert.That(paramA.Type, Is.EqualTo("string"));
            Assert.That(paramA.AllowedValues, Is.EquivalentTo(new[] { "A", "B", "C" }));

            var paramb = def.ParameterDefinition["paramb"];
            Assert.That(paramb.Type, Is.EqualTo("datetime"));
        }

        [Test]
        public void Basic_Parsing_of_array_parameter()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[arrayParameters]]
    tags");
            Assert.That(def.ArrayParameterSection.ArrayParameters.Count, Is.EqualTo(1));
        }

        [Test]
        public void Basic_Parsing_of_array_parameter_followed_by_other_params()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[arrayParameters]]
    tags=te,tr,ty
[[parameters]]
	TargetDateStart=2019-01-01");
            Assert.That(def.ArrayParameterSection.ArrayParameters.Count, Is.EqualTo(1));
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
    name: TestQueryName
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    template/Product Backlog Item: pbix.docx
    template/Bug: bugaa.docx
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.Name, Is.EqualTo("TestQueryName"));
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
    limit: 1
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.GetTemplateForWorkItem("Product Backlog Item"), Is.EqualTo("pbix.docx"));
            Assert.That(querySection.GetTemplateForWorkItem("Task"), Is.EqualTo(null));
            Assert.That(querySection.Limit, Is.EqualTo(1));
        }

        [Test]
        public void Query_section_parse_filter_work_item_by_type()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    template/Product Backlog Item: pbix.docx
    template/Bug: bugaa.docx
    limit: 1
    workItemTypes: Product Backlog Item,Feature
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.GetTemplateForWorkItem("Product Backlog Item"), Is.EqualTo("pbix.docx"));
            Assert.That(querySection.GetTemplateForWorkItem("Task"), Is.EqualTo(null));
            Assert.That(querySection.WorkItemTypes, Is.EquivalentTo(new[] { "Product Backlog Item", "Feature" }));
        }

        [Test]
        public void Query_section_without_limit_default_to_max_value()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.Limit, Is.EqualTo(Int32.MaxValue));
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
            Assert.That(def.ParameterSection.Parameters.Keys, Is.EquivalentTo(new[] { "parama", "paramb" }));
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

        [Test]
        public void Valid_query_iteration_parameter()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'""
    tableTemplate: table.docx
    repeatForEachIteration: true
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.RepeatForEachIteration, Is.True);
        }

        [Test]
        public void Valid_parametric_query()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT
        * 
        FROM workitems
        WHERE
            [System.TeamProject] = @project
            AND [System.WorkItemType] = 'Product Backlog Item'
            AND [System.IterationPath] = '{iterationPath}'""
    parameterSet: iterationPath=Zoalord Insurance\Release 1\Sprint 4
    parameterSet: iterationPath=Zoalord Insurance\Release 1\Sprint 5  
    parameterSet: iterationPath=Zoalord Insurance\Release 1\Sprint 6
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.QueryParameters.Count, Is.EqualTo(3));
            Assert.That(querySection.QueryParameters[0]["iterationPath"], Is.EqualTo(@"Zoalord Insurance\Release 1\Sprint 4"));
            Assert.That(querySection.QueryParameters[1]["iterationPath"], Is.EqualTo(@"Zoalord Insurance\Release 1\Sprint 5"));
            Assert.That(querySection.QueryParameters[2]["iterationPath"], Is.EqualTo(@"Zoalord Insurance\Release 1\Sprint 6"));
        }

        [Test]
        public void Multiline_query()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
query: ""SELECT
    * 
    FROM workitems
    WHERE
        [System.TeamProject] = @project
        AND [System.WorkItemType] = 'Product Backlog Item'
        AND [System.IterationPath] = '{iterationPath}'""
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.Query.Contains("AND [System.IterationPath] = '{iterationPath}'"));
            Assert.That(querySection.Query.Contains("[System.TeamProject] = @project"));
        }

        [Test]
        public void Multiline_query_complex()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT
        *
        FROM
            workitems
        WHERE
            [System.WorkItemType] = 'Feature' AND
            [Microsoft.VSTS.Scheduling.TargetDate] >= '{TargetDateStart}' AND
            [Microsoft.VSTS.Scheduling.TargetDate] <= '{TargetDateEnd}' AND
            NOT[System.Tags] CONTAINS 'outofrelease' AND

            NOT[System.Tags] CONTAINS 'Report' AND
            [System.TeamProject] = '{teamProjectName}'
          AND(
            [System.IterationPath] = '{Iteration1}'
            OR [System.IterationPath] = '{Iteration2}'
            OR [System.IterationPath] = '{Iteration3}'
            OR [System.IterationPath] = '{Iteration4}'
            Or [System.IterationPath] = '{Iteration5}'
          )""
    tableTemplate: 2_table.docx
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.Query.Contains("Or [System.IterationPath] = '{Iteration5}'"));
        }

        [Test]
        public void Query_with_hierarchy_mode()
        {
            var sut = new ConfigurationParser();
            TemplateDefinition def = sut.ParseTemplateDefinition(
@"[[query]]
    query: ""SELECT
    [System.Id],
    [System.WorkItemType],
    [System.Title],
    [System.AssignedTo],
    [System.State],
    [System.Tags]
FROM workitemLinks
WHERE
    (
        [Source].[System.TeamProject] = @project
        AND [Source].[System.WorkItemType] = 'Feature'
        AND [Source].[Microsoft.VSTS.Scheduling.TargetDate] < '2002-01-01T00:00:00.0000000'
        AND [Source].[Microsoft.VSTS.Scheduling.TargetDate] > '2000-02-02T00:00:00.0000000'
    )
    AND (
        [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward'
    )
    AND (
        [Target].[System.TeamProject] = @project
        AND [Target].[System.WorkItemType] <> ''
    )
MODE (Recursive)""
    hierarchyMode: task,feature,requirement,epic
");
            var querySection = def.AllSections.Single() as QuerySection;
            Assert.That(querySection.HierarchyMode, Is.EquivalentTo(new[] { "task", "feature", "requirement", "epic" }));
        }
    }
}
