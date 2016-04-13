using System;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Common;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class View : BaseModel, IEquatable<View>, IUrlProvider
    {
        #region Private Members
        private string _schemaXml = string.Empty;
        private XElement _node = null;
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets a value that specifies the XML Schema representing the View type.
        /// </summary>
        public string SchemaXml
        {
            get
            {
                return this._schemaXml;
            }
            set 
            { 
                if( value != this._schemaXml )
                {
                    this._schemaXml = value;
                    EnsureInitNode(true);
                }
            }
        }

        public XElement XmlNode
        {
            get
            {
                EnsureInitNode(false);
                return this._node;
            }
        }

        private void EnsureInitNode(bool force)
        {
            if (force || (null == this._node))
            {
                if (string.IsNullOrEmpty(this._schemaXml))
                {
                    this._node = null;
                }
                else
                {
                    this._node = XElement.Parse(this.SchemaXml);
                }
            }
        }

        public string PageUrl { get; set; }

        public string GetAttributeValue( string attrName )
        {
            string result = "";
            var node = this.XmlNode;
            if (null != node)
            {
                var attr = node.Attribute(attrName);
                if( null != attr )
                {
                    result = attr.Value;
                }
            }
            return result;
        }
        
        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            XElement element = PrepareViewForCompare(this.SchemaXml);
            return (element != null ? element.ToString().GetHashCode() : 0);
        }

        public string GetUrl()
        {
            return PageUrl;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is View))
            {
                return (false);
            }
            return (Equals((View)obj));
        }

        public bool Equals(View other)
        {
            if (other == null)
            {
                return (false);
            }

            XElement currentXml = PrepareViewForCompare(this.SchemaXml);
            XElement otherXml = PrepareViewForCompare(other.SchemaXml);
            return (XNode.DeepEquals(currentXml, otherXml));
        }

        private XElement PrepareViewForCompare(string schemaXML)
        {
            XElement element = XElement.Parse(schemaXML);
            if (element.Attribute("Name") != null)
            {
                Guid nameGuid = Guid.Empty;
                if (Guid.TryParse(element.Attribute("Name").Value, out nameGuid))
                {
                    // Temporary remove guid
                    element.Attribute("Name").Remove();
                }
            }

            //MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" 
            if (element.Attribute("ImageUrl") != null)
            {
                var index = element.Attribute("ImageUrl").Value.IndexOf("rev=", StringComparison.InvariantCultureIgnoreCase);

                if (index > -1)
                {
                    // Remove ?rev=23 in url
                    Regex regex = new Regex("\\?rev=([0-9])\\w+");
                    element.SetAttributeValue("ImageUrl", regex.Replace(element.Attribute("ImageUrl").Value, ""));
                }
            }

            string[] attrToRemove = new string[]
            {
                "Url",
                "MobileView",
                "MobileDefaultView",
                "Toolbar",
                "XslLink"
            };
            string[] elementsToDelete = new string[]
            {
                "Aggregations"
            };
            foreach (string attrName in attrToRemove)
            {
                XAttribute attr = element.Attribute(attrName);
                if (null != attr)
                {
                    attr.Remove();
                }
            }            
            foreach (string elName in elementsToDelete)
            {
                var child = element.Elements("Aggregations");
                if (null != child)
                {
                    child.Remove();
                }
            }

            return element;
        }

        #endregion
    }
}
