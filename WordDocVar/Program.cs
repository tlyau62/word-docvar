using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Engine;
using OpenXMLTemplates.Variables;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace WordDocVar
{
    class Program
    {
        static void Main(string[] args)
        {
            var engine = new DefaultOpenXmlTemplateEngine();
            var input = @"C:\Users\tyautl\Git\WordDocVar\WordDocVar\Resources\c.docx";
            var output = @"C:\Users\tyautl\Git\WordDocVar\WordDocVar\Resources\c2.docx";

            engine.RegisterReplacer(new HtmlControlReplacer());

            using (var doc = new TemplateDocument(input))
            {
                var src = new VariableSource(@"{ ""companyAddr"": ""4321"", ""content"": ""<p>test</p><p>testte<b>sascacs</b>t</p><ul><li>object</li><li>object2</li></ul>""}");

                engine.ReplaceAll(doc, src);

                doc.SaveAs(output);
            }

            Console.WriteLine("end");
        }
    }
}
