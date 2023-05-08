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
using AngleSharp.Xml.Parser;
using AngleSharp.Xml;
using AngleSharp.Html.Parser;

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
            var xml = new HtmlParser().ParseDocument(shtml).ToXml();
            var xe = XElement.Parse(xml);
            var wml = OpenXmlPowerTools.HtmlToWmlConverter.ConvertHtmlToWml("", "", "", xe, OpenXmlPowerTools.HtmlToWmlConverter.GetDefaultSettings());

            return ToOpenXmlElement(wml.MainDocumentPart);
        }
    }
}
