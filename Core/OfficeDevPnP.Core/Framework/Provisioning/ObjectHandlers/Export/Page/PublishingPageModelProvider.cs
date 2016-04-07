using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.Page
{
    internal class PublishingPageModelProvider :
        PageProviderBase,
        IPageModelProvider
    {
        public PublishingPageModelProvider(string homePageUrl, Web web, TokenParser parser) :
            base(homePageUrl, web, parser)
        {            
        }

        public override void AddPage(ListItem item, ProvisioningTemplate template)
        {
            string url = GetUrl(item, true);
            if (null == template.Pages.Find((p) => string.Equals(p.Url, url, System.StringComparison.OrdinalIgnoreCase)))
            {
                var fieldValues = item.FieldValues;

                var pageLayoutUrl = string.Empty;
                if (fieldValues.ContainsKey("PublishingPageLayout"))
                {
                    pageLayoutUrl = fieldValues["PublishingPageLayout"] == null ? "" : (fieldValues["PublishingPageLayout"] as FieldUrlValue).Url;
                }
                pageLayoutUrl = TokenParser.TokenizeUrl(this.Web, pageLayoutUrl);

                string html = "";
                if (fieldValues.ContainsKey("PublishingPageContent"))
                {
                    html = fieldValues["PublishingPageContent"] == null ? " " : fieldValues["PublishingPageContent"].ToString();
                }

                var title = fieldValues["Title"] == null ? "" : fieldValues["Title"].ToString();

                bool needToOverwrite = NeedOverride(item);
                var webParts = GetWebParts(item);
                bool isHomePage = IsWelcomePage(item);
                PublishingPage page = new PublishingPage(url, title, html, pageLayoutUrl, needToOverwrite, webParts, isHomePage);
                template.Pages.Add(page);
            }
        }
    }
}
