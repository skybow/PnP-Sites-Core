using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V2;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V3;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    class WebPartTokenizeManager
    {
        public static IWebPartTokenizer GetWebPartTokenizer(string xml) {
            if (WebPartsModelProvider.IsV3FormatXml(xml))
            {
                return GetV3WebPartTokenizer(xml);
            }
            else 
            {
                return GetV2WebPartTokenizer(xml);
            }
        }

        private static IWebPartTokenizer GetV3WebPartTokenizer(string xml) {
            XElement webPartXml = XElement.Parse(xml);
            //var webPartTypeNode = webPartXml.XPathSelectElement("/webParts/webPart/metaData/type");
            string webPartTypeWithAssebly = webPartXml.Descendants().FirstOrDefault(n => n.Name.LocalName.Equals("type", StringComparison.InvariantCultureIgnoreCase)).Attribute("name").Value;
            var webPartType = webPartTypeWithAssebly.Split(',')[0];
            return V3WebPartTokenizerManager.GetWebPartTokenizer(webPartType);
        }

        private static IWebPartTokenizer GetV2WebPartTokenizer(string xml)
        {
            XElement webPartXml = XElement.Parse(xml);
            var webPartTypeNode = webPartXml.Nodes().FirstOrDefault(n => (n as XElement).Name.LocalName.Equals("TypeName", StringComparison.InvariantCultureIgnoreCase));
            string webPartType = (webPartTypeNode as XElement).Value;
            return V2WebPartTokenizerManager.GetWebPartTokenizer(webPartType);
        }
    }
}
