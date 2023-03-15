using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Engine;
using OpenXMLTemplates.Variables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Xml.Linq;

namespace WordDocVar
{
    class Program
    {
        static IDictionary<string, Assembly> Dictionary = AppDomain.CurrentDomain.GetAssemblies().ToDictionary(e => e.FullName);

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return null;
        }

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            var content = File.ReadAllText(@"C:\Users\tyautl\Git\WordDocVar\WordDocVar\Resources\test.html");
            var engine = new DefaultOpenXmlTemplateEngine();
            var input = @"C:\Users\tyautl\Git\WordDocVar\WordDocVar\Resources\memo.docx";
            var output = @"C:\Users\tyautl\Git\WordDocVar\WordDocVar\Resources\memo2.docx";

            engine.RegisterReplacer(new HtmlControlReplacer());

            using (var doc = new TemplateDocument(input))
            {
                //var src = new VariableSource(@$"{{ 
                //    ""from"": ""Tim"", 
                //    ""to"": ""Benny"",
                //    ""ourRefFolio"": ""ef(1)"",
                //    ""ourRefNo"": ""esh001"",
                //    ""senderTel"": ""3345678"",
                //    ""senderFax"": ""12345678"",
                //    ""dated"": ""01/02/2022"",
                //    ""attn"": ""Tim, Benny"",
                //    ""yourRefFolio"": ""ef(2)"",
                //    ""yourRefNo"": ""esh002"",
                //    ""yourDated"": ""02/02/2022"",
                //    ""recipientFax"": ""987654321"",
                //    ""title"": ""A memo title"",
                //    ""content"": ""{content}"",
                //    ""contentToApplicant"": ""<p>Byebye, I am <b>Tim</b></p>"",
                //}}");

                var src = new VariableSource(new Dictionary<string, string>()
                {
                    { "from", "Tim" }, 
                    { "to", "Benny" },
                    { "ourRefFolio", "ef(1)" },
                    { "ourRefNo", "esh001" },
                    { "senderTel", "3345678" },
                    { "senderFax", "12345678" },
                    { "dated", "01 / 02 / 2022" },
                    { "attn", "Tim, Benny" },
                    { "yourRefFolio", "ef(2)" },
                    { "yourRefNo", "esh002" },
                    { "yourDated", "02 / 02 / 2022" },
                    { "recipientFax", "987654321" },
                    { "title", "A memo title" },
                    { "contentToApplicant", "<p>Byebye, I am<b>Tim</b></p>" },
                    { "content", content}
                });

                engine.ReplaceAll(doc, src);

                doc.SaveAs(output);
            }

            Console.WriteLine("end");
        }
    }
}
