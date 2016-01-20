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
    internal class PublishingPageModelProvider : IPageModelProvider
    {
        protected string HomePageUrl { get; private set; }
        protected Web Web { get; private set; }
        protected WebPartsModelProvider Provider;
        public PublishingPageModelProvider(string homePageUrl, Web web)
        {
            HomePageUrl = homePageUrl;
            this.Web = web;
            Provider = new WebPartsModelProvider(web);
        }
        public ContentPage GetPage(ListItem item)
        {
            var html = string.Empty;
            var fieldValues = item.FieldValues;
            var title = fieldValues["Title"] == null ? string.Empty : fieldValues["Title"].ToString();
            if (fieldValues.ContainsKey("PublishingPageContent"))
            {
                html = fieldValues["PublishingPageContent"] == null ? " " : fieldValues["PublishingPageContent"].ToString();
            }

            var pageLayoutUrl = string.Empty;
            if (fieldValues.ContainsKey("PublishingPageLayout"))
            {
                pageLayoutUrl = fieldValues["PublishingPageLayout"] == null ? "" : (fieldValues["PublishingPageLayout"] as FieldUrlValue).Url;
            }

            var siteCollectionContext = Web.Context.GetSiteCollectionContext();
            pageLayoutUrl = pageLayoutUrl.Replace(siteCollectionContext.Url, string.Empty);

            var url = fieldValues["FileRef"].ToString();
            var isHomePage = HomePageUrl.Equals(url);
            var needToOverwrite = item.File.Versions.Any();

            var webParts = Provider.Retrieve(url);
            url = url.Replace(Web.RootFolder.ServerRelativeUrl, "~site/");
            return new PublishingPage(url, title, html, pageLayoutUrl, needToOverwrite, webParts, isHomePage);
        }
    }
}
