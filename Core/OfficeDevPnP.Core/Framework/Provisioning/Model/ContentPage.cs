using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class ContentPage : Page
    {
        public string Html { get; set; }

        public ContentPage(string url, string html, bool overwrite, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null)
            : base(url, overwrite, WikiPageLayout.OneColumn, webParts, welcomePage, security)
        {
            Html = html;
        }
    }
}
