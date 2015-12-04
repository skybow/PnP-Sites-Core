using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Entities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    internal class WebPartsEntityProvider
    {
        protected Web Web { get; set; }
        protected string PageUrl { get; set; }

        public WebPartsEntityProvider(Web web, string pageUrl)
        {
            Web = web;
            PageUrl = pageUrl;
        }

        private static readonly Regex GetWebPartXmlReqex = new Regex(@"<WebPart (.*?)<\/WebPart>", RegexOptions.Singleline);
        private static readonly Regex ZoneRegEx = new Regex(@"(?<=ZoneID>).*?(?=<\/ZoneID>)");
        private static readonly Regex PartOrderReqEx = new Regex(@"(?<=PartOrder>).*?(?=<\/PartOrder>)");

        public List<WebPartEntity> Retrieve(string xml, string pageUrl)
        {
            var result = new List<WebPartEntity>();
            if (string.IsNullOrEmpty(xml)) return result;
            var maches = GetWebPartXmlReqex.Matches(xml);

            foreach (var match in maches)
            {
                var webPartXml = match.ToString();
                var zone = this.GetZone(webPartXml);
                var zoneIndex = this.GetZoneIndex(webPartXml);
                webPartXml = this.WrapToV3Format(webPartXml);
                
                var entity = new WebPartEntity
                {
                    WebPartXml = webPartXml,
                    WebPartIndex = zoneIndex,
                    WebPartZone = zone
                };

                result.Add(entity);
            }

            return result;
        }

        private int GetZoneIndexWithServiceCall(Guid webPartId)
        {
            var page = Web.GetFileByServerRelativeUrl(PageUrl);
            var manager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webParts = manager.WebParts;
            var context = Web.Context;
            context.Load(webParts);
            context.ExecuteQueryRetry();
            var webPart = webParts.GetById(webPartId).WebPart;
            if (webPart == null) return 0;
            context.Load(webPart);
            context.ExecuteQueryRetry();
            return webPart.ZoneIndex;
        }

        private string GetZone(string webPartXml)
        {
            var value = ZoneRegEx.Match(webPartXml).Value;
            return string.IsNullOrEmpty(value) ? "Main" : value;
        }

        private int GetZoneIndex(string webPartXml)
        {
            var partOrderMatch = PartOrderReqEx.Match(webPartXml).Value;
            if (!string.IsNullOrEmpty(partOrderMatch)) return int.Parse(partOrderMatch);
            var webPartId = this.GetWebPartId(webPartXml);
            return this.GetZoneIndexWithServiceCall(webPartId);
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
