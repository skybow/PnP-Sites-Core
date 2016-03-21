using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V3
{
    class V3DefaultWebPartTokenizer : IWebPartTokenizer
    {
        protected List<string> NodesToSkip = new List<string>();

        public V3DefaultWebPartTokenizer() { 
        }

        public string Tokenize(string xml, TokenParser parser) {

            XElement webPartXml = XElement.Parse(xml);
            var webPartProperties = webPartXml.Descendants().FirstOrDefault(n => n.Name.LocalName.Equals("properties", StringComparison.InvariantCultureIgnoreCase));
            var nodes = webPartProperties.Nodes();
            foreach (var node in nodes)
            {
                var element = node as XElement;
                if (!SkipTokenization(element.Name.LocalName))
                {
                    element.Value = parser.TokenizeString(element.Value);
                }

            }

            return webPartXml.ToString();
        }

        protected bool SkipTokenization(string NodeName)
        {
            return NodesToSkip.Any(n => n.Equals(NodeName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
