using DocumentFormat.OpenXml;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net;
using Ganss.Xss;
using Microsoft.Security.Application;
using System.Xml;
using System.Text.RegularExpressions;
using System.Web;

namespace WordDocVar
{
    public class HtmlControlReplacer : ControlReplacer
    {
        public override string TagName => "html";

        protected override OpenXmlExtensions.ContentControlType ContentControlTypeRestriction => OpenXmlExtensions.ContentControlType.RichText;

        protected override string ProcessControl(string variableIdentifier, IVariableSource variableSource, ContentControl contentControl, List<string> otherParameters)
        {
            return variableSource.GetVariable<string>(variableIdentifier);
        }

        protected override void OnReplaced(ContentControl e)
        {
            var html = e.SdtElement.InnerText;
            var oml = ConvertHtmlToÓml(html) as Document;
            var nodes = oml.Body.Elements()
                .SkipLast(1)
                .Select(n => n.CloneNode(true));

            e.SdtElement.RemoveAllChildren();

            foreach (var node in nodes)
            {
                e.SdtElement.AppendChild(node);
            }

            base.OnReplaced(e);
        }

        private OpenXmlElement ToOpenXmlElement(XElement element)
        {
            // Write XElement to MemoryStream.
            using var stream = new MemoryStream();
            element.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);

            // Read OpenXmlElement from MemoryStream.
            using OpenXmlReader reader = OpenXmlReader.Create(stream);
            reader.Read();
            return reader.LoadCurrentElement();
        }

        private OpenXmlElement ConvertHtmlToÓml(string html)
        {
            var htmlSanitizer = new HtmlSanitizer();
            var shtml = htmlSanitizer.Sanitize(html);
            var xml = Regex.Replace(shtml, @"&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-fA-F]{1,6});", m => EscapeXml(HttpUtility.HtmlDecode(m.Value)));
            var wrap = $"<html><body>{xml}</body></html>";
            var xe = XElement.Parse(wrap);
            var wml = OpenXmlPowerTools.HtmlToWmlConverter.ConvertHtmlToWml("", "", "", xe, OpenXmlPowerTools.HtmlToWmlConverter.GetDefaultSettings());

            return ToOpenXmlElement(wml.MainDocumentPart);
        }

        /**
         * https://stackoverflow.com/questions/22906722/how-to-encode-special-characters-in-xml
         */
        private string EscapeXml(string s)
        {
            string toxml = s;

            if (!string.IsNullOrEmpty(toxml))
            {
                // replace literal values with entities
                toxml = toxml.Replace("&", "&amp;");
                toxml = toxml.Replace("'", "&apos;");
                toxml = toxml.Replace("\"", "&quot;");
                toxml = toxml.Replace(">", "&gt;");
                toxml = toxml.Replace("<", "&lt;");
            }

            return toxml;
        }
    }
}
