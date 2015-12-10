using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Common;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class Form : IUrlProvider
    {
        public PageType FormType { get; set; }
        public string ServerRelativeUrl { get; set; }
        public bool IsDefault { get; set; }
        public string GetUrl()
        {
            return ServerRelativeUrl;
        }
    }
}
