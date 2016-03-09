using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class PublishingPage: ContentPage
    {
        public string PageTitle { get; set; }
        public string PageLayoutUrl { get; set; }

        public PublishingPage(string url, string title,  string html, string layoutUrl, bool overwrite, IEnumerable<WebPart> webParts, bool welcomePage = false, ObjectSecurity security = null, Dictionary<string, string> fields = null)
            : base(url, html, overwrite, webParts, welcomePage, security, fields)
        {
            PageTitle = title;
            PageLayoutUrl = layoutUrl;
        }
    }
}
