using System.Linq;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts.V2
{
    public class V2WebPartPropertiesCleaner :
        WebPartPropertiesCleanerBase
    {
        private string[] _skipProperties = null;

        public V2WebPartPropertiesCleaner(string[] skipProperties)
        {
            this._skipProperties = skipProperties;
        }

        public override string XElementPropertiesPath
        {
            get { return null; }
        }

        protected override bool IsPropertyDefault(XElement xmlProp)
        {
            bool result = false;

            string propName = xmlProp.Name.LocalName;
            string propValue = xmlProp.Value;

            if (string.IsNullOrEmpty(propName) || string.IsNullOrEmpty(propValue) ||
                this._skipProperties.Contains(propName))
            {
                result = true;
            }
            return result;
        }
    }
}
