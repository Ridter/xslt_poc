using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Xml;
namespace ApplyXSLTToXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string XMLFilePath = @"http://evi1cg.me/scripts/example.xml";
            string XSLTFilePath = @"http://evi1cg.me/scripts/calc.xslt";
            try
            {
                XsltSettings xslt_settings = new XsltSettings(false, true);
                xslt_settings.EnableScript = true;
                XslCompiledTransform xslt = new XslCompiledTransform();
                XmlUrlResolver resolver = new XmlUrlResolver();
                // Load documents 
                xslt.Load(XSLTFilePath, xslt_settings, resolver);
                XPathDocument xmlPathDoc = new XPathDocument(XMLFilePath);
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.OmitXmlDeclaration = true;
                XmlWriter writer = XmlWriter.Create("output.xml", settings);
                xslt.Transform(xmlPathDoc,writer);
                writer.Close();

            }
            catch (Exception e)
            {

                Console.WriteLine("Error:");
                Console.WriteLine(e.Message);

            }
        }
    }
}