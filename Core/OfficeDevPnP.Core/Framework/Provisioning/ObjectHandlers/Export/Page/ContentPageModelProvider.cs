using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.Page
{
    internal class ContentPageModelProvider : IPageModelProvider
    {
        protected string HomePageUrl { get; private set; }
        protected Web Web { get; private set; }
        protected WebPartsModelProvider Provider;
        public ContentPageModelProvider(string homePageUrl, Web web)
        {
            HomePageUrl = homePageUrl;
            this.Web = web;
            Provider = new WebPartsModelProvider(web);
        }

        public ContentPage GetPage(ListItem item)
        {
            var html = string.Empty;
            var fieldValues = item.FieldValues;
            if (fieldValues.ContainsKey("WikiField"))
            {
                html = fieldValues["WikiField"] == null ? " " : fieldValues["WikiField"].ToString();
            }
            var url = fieldValues["FileRef"].ToString();
            var isHomePage = HomePageUrl.Equals(url);
            var needToOverwrite = item.File.Versions.Any();

            var webParts = Provider.Retrieve(url);
            url = this.TokenizeUrl(url);

            return new ContentPage(url, html, needToOverwrite, webParts, isHomePage);
        }

        private string TokenizeUrl(string url)
        {
            url = url.Replace(Web.Url, "{site}/");
            url = url.Replace(Web.RootFolder.ServerRelativeUrl, "{site}/");

            var context = Web.Context as ClientContext;
            var site = context.Site;
            site.EnsureProperties(s => s.Url, s => s.ServerRelativeUrl);

            url = url.Replace(site.Url, "{sitecollection}/");
            if (site.ServerRelativeUrl == "/")
            {
                if (url.StartsWith("/"))
                {
                    url = "{sitecollection}" + url;
                }
            }
            else
            {
                url = url.Replace(site.ServerRelativeUrl, "{sitecollection}/");
            }

            return url;
        }
    }
}
