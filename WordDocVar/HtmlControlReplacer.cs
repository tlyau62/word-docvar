using DocumentFormat.OpenXml;
using OpenXMLTemplates;
using OpenXMLTemplates.ControlReplacers;
using OpenXMLTemplates.Documents;
using OpenXMLTemplates.Variables;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

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

            foreach (var node in nodes) { 
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
            var xe = XElement.Parse($"<html><body>{html}</body></html>");
            var wml = OpenXmlPowerTools.HtmlToWmlConverter.ConvertHtmlToWml("", "", "", xe, OpenXmlPowerTools.HtmlToWmlConverter.GetDefaultSettings());

            return ToOpenXmlElement(wml.MainDocumentPart);
        }
    }
}
