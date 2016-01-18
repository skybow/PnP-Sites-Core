using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class ContentPage : Page
    {
        public string PageTitle { get; set; }
        public string Html { get; set; }

        public ContentPage(string url, string title, string html, bool overwrite, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null)
            : base(url, overwrite, WikiPageLayout.OneColumn, webParts, welcomePage, security)
        {
            PageTitle = title;
            Html = html;
        }
    }
}
