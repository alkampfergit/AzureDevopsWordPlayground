using NUnit.Framework;
using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Templates.Parser;

namespace WordExporter.Tests.Templates.Parser
{
    [TestFixture]
    public class SingleSectionParserTests
    {
        [Test]
        public void Identify_basic_parameters_section()
        {
            var section = ConfigurationParser.SectionParser.Parse(@"[[parameters]]
section");
            Assert.That(section, Is.InstanceOf<ParameterSection>());
        }

        [Test]
        public void Unknown_section_Will_throw()
        {
            Assert.Throws<NotSupportedException>(() => ConfigurationParser.SectionParser.Parse("[[unknown section]]"));
        }

        [Test]
        public void Basic_KeyValue()
        {
            var keyValue = ConfigurationParser.KeyValue.Parse("key: value");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value"));
        }

        [Test]
        public void Basic_Multiple_KeyValue()
        {
            var keyValues = ConfigurationParser.KeyValueList.Parse("key: value\nkey2: value2").ToArray();
            Assert.That(keyValues.Length, Is.EqualTo(2));
            Assert.That(keyValues[0].Key, Is.EqualTo("key"));
            Assert.That(keyValues[0].Value, Is.EqualTo("value"));
            Assert.That(keyValues[1].Key, Is.EqualTo("key2"));
            Assert.That(keyValues[1].Value, Is.EqualTo("value2"));
        }

        [Test]
        public void Basic_KeyValue_resilient_to_spaces()
        {
            var keyValue = ConfigurationParser.KeyValue.Parse("key  : value  ");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value"));
        }

        [Test]
        public void SingleLine_with_semicolon()
        {
            var keyValue = ConfigurationParser.KeyValue.Parse("key: value:with:semicolon");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value:with:semicolon"));
        }

        [Test]
        public void MultiLine_key_value()
        {
            var keyValue = ConfigurationParser.MultiLineKeyValue.Parse(@"key: ""value

value2

value3""");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value\r\n\r\nvalue2\r\n\r\nvalue3"));
        }

        [Test]
        public void MultiLine_key_value_can_contain_square_bracket()
        {
            var keyValue = ConfigurationParser.MultiLineKeyValue.Parse(@"key: ""value
[value2]
value3""");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value\r\n[value2]\r\nvalue3"));
        }

        [Test]
        public void MultiLine_key_value_can_contain_semicolon()
        {
            var keyValue = ConfigurationParser.MultiLineKeyValue.Parse(@"key: ""value
value:[with]:semicolon
:value3""");
            Assert.That(keyValue.Key, Is.EqualTo("key"));
            Assert.That(keyValue.Value, Is.EqualTo("value\r\nvalue:[with]:semicolon\r\n:value3"));
        }
    }
}
