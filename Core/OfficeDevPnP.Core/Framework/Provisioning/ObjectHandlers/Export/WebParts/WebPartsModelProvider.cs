using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using ModelWebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    internal class WebPartsModelProvider
    {
        protected Web Web { get; set; }

        private static readonly Regex GetWebPartXmlReqex = new Regex(@"<WebPart (.*?)<\/WebPart>", RegexOptions.Singleline);
        private static readonly Regex ZoneRegEx = new Regex(@"(?<=ZoneID>).*?(?=<\/ZoneID>)", RegexOptions.Singleline);
        private static readonly Regex WebPartIdEx = new Regex(@"(?<=<ID>).*?(?=<\/ID>)", RegexOptions.Singleline);

        public WebPartsModelProvider(Web web)
        {
            Web = web;
        }

        public List<ModelWebPart> Retrieve(string pageUrl, TokenParser parser)
        {
            if (parser == null) { 
                parser = new TokenParser(this.Web, new Model.ProvisioningTemplate());
            }
            var xml = Web.GetWebPartsXml(pageUrl);
            var pageContent = Web.GetPageContent(pageUrl);
            var result = new List<ModelWebPart>();
            if (string.IsNullOrEmpty(xml)) return result;
            var maches = GetWebPartXmlReqex.Matches(xml);

            var webPartDefinitions = this.GetWebPartDefinitionsWithServiceCall(pageUrl);

            foreach (var match in maches)
            {
                var webPartXml = match.ToString();
                var zone = this.GetZone(webPartXml);
                var wpId = this.GetWebPartId(webPartXml);

                var definition = webPartDefinitions.FirstOrDefault(d => d.Id == wpId);
                var webPart = definition.WebPart;
                webPartXml = this.WrapToV3Format(webPartXml);
                var pcLower = pageContent.ToLower();
                //TODO: refactor getting webpartId2 make separate method, probably use regex or another approach
                var contentBoxIndex = pcLower.IndexOf("<div id=\"contentbox\"");
                var indexOfIdStartIndex = contentBoxIndex != -1
                    ? pcLower.IndexOf(wpId.ToString().ToLower(), contentBoxIndex)
                    : -1;
                var indexOfId = indexOfIdStartIndex != -1 ? pcLower.IndexOf("webpartid2", indexOfIdStartIndex) : -1;
                var wpExportId = definition.Id;
                var wpControlId = GetWebPartControlId(webPartXml);
                if (indexOfId != -1 && string.IsNullOrEmpty(wpControlId))
                {
                    var wpId2 = pageContent.Substring(indexOfId + "webpartid2=\"".Length, 36);
                    wpExportId = Guid.Parse(wpId2);
                }

                webPartXml = this.SetWebPartIdToXml(wpExportId, webPartXml);
                webPartXml = this.TokenizeWebPartXml(webPartXml, parser);
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

        public static string GetWebPartControlId(string webPartXml)
        {
            var value = WebPartIdEx.Match(webPartXml).Value;
            return value;
        }

        private string SetWebPartIdToXml(Guid id, string xml)
        {
            var element = XElement.Parse(xml);
            element.SetAttributeValue("webpartid", id);
            return element.ToString();
        }

        private string TokenizeWebPartXml(string xml, TokenParser parser)
        {
            //TODO tokenize specific properties for specific webparts
            //Todo fix server relateive url "/" problem
            var tokenizer = WebPartTokenizeManager.GetWebPartTokenizer(xml);
            return tokenizer.Tokenize(xml, parser);
        }

        private List<WebPartDefinition> GetWebPartDefinitionsWithServiceCall(string pageUrl)
        {
            var definitions = new List<WebPartDefinition>();
            var page = Web.GetFileByServerRelativeUrl(pageUrl);
            var manager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webParts = manager.WebParts;
            var context = Web.Context;
            context.Load(webParts, wp => wp.Include(x => x.Id, x => x.WebPart.Title, x => x.WebPart.ZoneIndex));
            context.ExecuteQueryRetry();

            foreach (var definition in webParts)
            {
                definitions.Add(definition);
            }

            return definitions;
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
            if (!IsV3FormatXml(webPartxXml))
                return webPartxXml;
            var getWebPartXmlReqex = new Regex(@"<webPart (.*?)<\/webPart>", RegexOptions.Singleline);
            webPartxXml = getWebPartXmlReqex.Match(webPartxXml).Value;
            return string.Format("<webParts>{0}</webParts>", webPartxXml);
        }

        public static bool IsV3FormatXml(string xml)
        {
            return xml.IndexOf("http://schemas.microsoft.com/WebPart/v3", StringComparison.OrdinalIgnoreCase) != -1;
        }

        public static bool IsWebPartDefault(ModelWebPart wp)
        {
            var wpcomparer = WebPartSchemaComparer.CreateTypedComparer(wp);
            bool result = (null != wpcomparer) && wpcomparer.IsDefaultWebPart(wp);

            return result;
        }
    }
}
