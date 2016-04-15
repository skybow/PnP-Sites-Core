using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    public abstract class WebPartPropertiesCleanerBase
    {
        public abstract string XElementPropertiesPath { get; }

        public virtual void CleanDefaultProperties(XElement xmlWebPart)
        {
            //Remove default properties

            XElement xmlProps = string.IsNullOrEmpty(this.XElementPropertiesPath) ?
                xmlWebPart : xmlWebPart.XPathSelectElement(this.XElementPropertiesPath);
            if (null != xmlProps)
            {
                var nodes = xmlProps.Nodes().ToList();
                for (var i = nodes.Count() - 1; i >= 0; i--)
                {
                    var xmlProp = nodes[i] as XElement;
                    if ((null != xmlProp) && IsPropertyDefault(xmlProp))
                    {
                        xmlProp.Remove();
                    }
                }
            }
        }

        protected virtual bool IsPropertyDefault(XElement xmlProperty)
        {
            return false;
        }
    }
}
