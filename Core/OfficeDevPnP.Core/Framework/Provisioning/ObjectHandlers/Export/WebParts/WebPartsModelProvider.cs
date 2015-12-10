using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using ModelWebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    internal class WebPartsModelProvider
    {
        protected Web Web { get; set; }
        protected string PageUrl { get; set; }

        public WebPartsModelProvider(Web web, string pageUrl)
        {
            Web = web;
            PageUrl = pageUrl;
        }

        private static readonly Regex GetWebPartXmlReqex = new Regex(@"<WebPart (.*?)<\/WebPart>", RegexOptions.Singleline);
        private static readonly Regex ZoneRegEx = new Regex(@"(?<=ZoneID>).*?(?=<\/ZoneID>)");

        public List<ModelWebPart> Retrieve(string xml, string pageUrl)
        {
            var result = new List<ModelWebPart>();
            if (string.IsNullOrEmpty(xml)) return result;
            var maches = GetWebPartXmlReqex.Matches(xml);

            foreach (var match in maches)
            {
                var webPartXml = match.ToString();
                var zone = this.GetZone(webPartXml);
                var webPart = this.GetWebPartWithServiceCall(this.GetWebPartId(webPartXml));
                webPartXml = this.WrapToV3Format(webPartXml);
                var entity = new ModelWebPart
                {
                    Contents = webPartXml,
                    Order = (uint) webPart.ZoneIndex,
                    Zone = zone,
                    Title = webPart.Title
                };

                result.Add(entity);
            }

            return result;
        }

        private WebPart GetWebPartWithServiceCall(Guid webPartId)
        {
            var page = Web.GetFileByServerRelativeUrl(PageUrl);
            var manager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webParts = manager.WebParts;
            var context = Web.Context;
            context.Load(webParts);
            context.ExecuteQueryRetry();
            var webPart = webParts.GetById(webPartId).WebPart;
            context.Load(webPart, x=>x.Title, x=>x.ZoneIndex);
            context.ExecuteQueryRetry();
            return webPart;
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
