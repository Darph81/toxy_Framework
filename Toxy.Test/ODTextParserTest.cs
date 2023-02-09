using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class ODTextParserTest
    {
        [Test]
        public void TestLoadODSFIle()
        {
            string path = TestDataSample.GetLibreOfficePath("testfile.ods");

            ParserContext context = new ParserContext(path);
            ITextParser parser = ParserFactory.CreateText(context);
            string text = parser.Parse();
            Assert.AreEqual("Here is some text to extract.\r\n", text);
        }

        [Test]
        public void TestLoadODTFIle()
        {
            string path = TestDataSample.GetLibreOfficePath("testfile.odt");

            ParserContext context = new ParserContext(path);
            ITextParser parser = ParserFactory.CreateText(context);
            string text = parser.Parse();
            Assert.AreEqual("Here is some text to extract.\r\n", text);
        }

        [Test]
        public void TestLoadODPFIle()
        {
            string path = TestDataSample.GetLibreOfficePath("testfile.odp");

            ParserContext context = new ParserContext(path);
            ITextParser parser = ParserFactory.CreateText(context);
            string text = parser.Parse();
            Assert.AreEqual("Hereissometexttoextract.\r\n", text);
        }
    }
}
