using System.Linq;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V3
{
    public class V3WebPartPropertiesCleaner :
        WebPartPropertiesCleanerBase
    {
        private string[] _skipProperties = null;

        public V3WebPartPropertiesCleaner(string[] skipProperties)
        {
            this._skipProperties = skipProperties;
        }

        public override string XElementPropertiesPath
        {
            get { return "//*[local-name() = 'properties']"; }
        }

        protected override bool IsPropertyDefault(XElement xmlProp)
        {
            bool result = false;

            XAttribute attrName = xmlProp.Attribute("name");
            if (null != attrName)
            {
                string propName = attrName.Value;
                string propValue = xmlProp.Value;

                if (string.IsNullOrEmpty(propName) || string.IsNullOrEmpty(propValue) ||
                    this._skipProperties.Contains(propName))
                {
                    result = true;
                }
            }

            return result;
        }
    }    
}
