using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.Page
{
    internal class ContentPageModelProvider :
        PageProviderBase,
        IPageModelProvider
    {
        public ContentPageModelProvider(string homePageUrl, Web web, TokenParser parser) :
            base(homePageUrl, web, parser)
        {            
        }

        public override void AddPage(ListItem item, ProvisioningTemplate template)
        {
            string url = GetUrl(item, true);
            if (null == template.Pages.Find((p) => string.Equals(p.Url, url, System.StringComparison.OrdinalIgnoreCase)))
            {
                var fieldValues = item.FieldValues;
                string html = "";
                if (fieldValues.ContainsKey("WikiField"))
                {
                    html = fieldValues["WikiField"] == null ? " " : fieldValues["WikiField"].ToString();
                }
                var title = fieldValues["Title"] == null ? "" : fieldValues["Title"].ToString();
                bool needToOverwrite = NeedOverride(item);
                var webParts = GetWebParts(item);
                bool isHomePage = IsWelcomePage(item);

                ContentPage page = new ContentPage(url, html, needToOverwrite, webParts, isHomePage);
                template.Pages.Add(page);
            }
        }
    }
}
