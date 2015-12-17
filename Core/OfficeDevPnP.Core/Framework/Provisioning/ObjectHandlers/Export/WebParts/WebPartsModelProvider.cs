using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using ModelWebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    internal class WebPartsModelProvider
    {
        protected Web Web { get; set; }

        private static readonly Regex GetWebPartXmlReqex = new Regex(@"<WebPart (.*?)<\/WebPart>", RegexOptions.Singleline);
        private static readonly Regex ZoneRegEx = new Regex(@"(?<=ZoneID>).*?(?=<\/ZoneID>)");

        public WebPartsModelProvider(Web web)
        {
            Web = web;
        }

        public List<ModelWebPart> Retrieve(string pageUrl)
        {
            var xml = Web.GetWebPartsXml(pageUrl);
            var result = new List<ModelWebPart>();
            if (string.IsNullOrEmpty(xml)) return result;
            xml = this.TokenizeXml(xml);
            var maches = GetWebPartXmlReqex.Matches(xml);

            foreach (var match in maches)
            {
                var webPartXml = match.ToString();
                var zone = this.GetZone(webPartXml);
                var definition = this.GetWebPartDefinitionWithServiceCall(this.GetWebPartId(webPartXml), pageUrl);
                var webPart = definition.WebPart;
                webPartXml = this.WrapToV3Format(webPartXml);
                webPartXml = this.SetWebPartIdToXml(definition.Id, webPartXml);

                var entity = new ModelWebPart
                {
                    Contents = webPartXml,
                    Order = (uint)webPart.ZoneIndex,
                    Zone = zone,
                    Title = webPart.Title
                };

                result.Add(entity);
            }

            return result;
        }

        private string SetWebPartIdToXml(Guid id, string xml)
        {
            var element = XElement.Parse(xml);
            element.SetAttributeValue("webpartid", id);
            var reader = element.CreateReader();
            reader.MoveToContent();
            return reader.ReadOuterXml();
        }

        private string TokenizeXml(string xml)
        {
            xml = xml.Replace(Web.ServerRelativeUrl, "~site");
            return xml.Replace(Web.Id.ToString(), "~siteid");
        }

        private WebPartDefinition GetWebPartDefinitionWithServiceCall(Guid webPartId, string pageUrl)
        {
            var page = Web.GetFileByServerRelativeUrl(pageUrl);
            var manager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webParts = manager.WebParts;
            var definition = webParts.GetById(webPartId);
            var context = Web.Context;
            context.Load(definition, x=>x.Id, x => x.WebPart.Title, x => x.WebPart.ZoneIndex);
            context.ExecuteQueryRetry();
            return definition;
        }

        private string GetZone(string webPartXml)
        {
            var value = ZoneRegEx.Match(webPartXml).Value;
            return string.IsNullOrEmpty(value) ? "Main" : value;
        }

        private Guid GetWebPartId(string webPartXml)
        {
            var stringToFind = "ID=\"";
            var index = webPartXml.IndexOf(stringToFind, StringComparison.Ordinal) + stringToFind.Length;
            var id = webPartXml.Substring(index, Guid.Empty.ToString().Length);
            return new Guid(id);
        }

        private string WrapToV3Format(string webPartxXml)
        {
            if (webPartxXml.IndexOf("http://schemas.microsoft.com/WebPart/v3", StringComparison.OrdinalIgnoreCase) == -1)
                return webPartxXml;
            var getWebPartXmlReqex = new Regex(@"<webPart (.*?)<\/webPart>", RegexOptions.Singleline);
            webPartxXml = getWebPartXmlReqex.Match(webPartxXml).Value;
            return string.Format("<webParts>{0}</webParts>", webPartxXml);
        }
    }
}
