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
    internal interface IPageModelProvider
    {
        void AddPage(ListItem item, ProvisioningTemplate template);
    }

    internal abstract class PageProviderBase :
        IPageModelProvider
    {
        protected string HomePageUrl { get; private set; }
        protected Web Web { get; private set; }
        protected WebPartsModelProvider Provider{ get; set; }
        protected TokenParser TokenParser { get; set; }

        public PageProviderBase(string homePageUrl, Web web, TokenParser parser)
        {
            this.HomePageUrl = homePageUrl;
            this.Web = web;
            this.Provider = new WebPartsModelProvider(web);
            this.TokenParser = parser;
        }

        public abstract void AddPage(ListItem item, ProvisioningTemplate template);

        protected virtual bool NeedOverride(ListItem item)
        {
            item.File.EnsureProperty(f => f.Versions);
            var needToOverwrite = item.File.Versions.Any();
            return needToOverwrite;
        }

        protected string GetUrl( ListItem item, bool tokenize )
        {
            string url = item.FieldValues["FileRef"].ToString();
            if (tokenize)
            {
                url = TokenParser.TokenizeUrl(this.Web, url);
            }
            return url;
        }

        protected List<WebPart> GetWebParts( ListItem item )
        {
            string url = GetUrl(item, false);
            var webParts = Provider.Retrieve(url, this.TokenParser );
            return webParts;
        }

        protected bool IsWelcomePage(ListItem item)
        {
            string url = GetUrl(item, false);
            bool result = url.Equals(this.HomePageUrl, StringComparison.OrdinalIgnoreCase);
            return result;
        }
    }
}
