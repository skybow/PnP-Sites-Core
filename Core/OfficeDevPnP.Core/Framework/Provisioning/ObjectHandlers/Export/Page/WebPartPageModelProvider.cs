using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;
using System;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;
using System.Web;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.File;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.Page
{
    internal class WebPartPageModelProvider:
        PageProviderBase,
        IPageModelProvider
    {
        public WebPartPageModelProvider(string homePageUrl, Web web, TokenParser parser) :
            base(homePageUrl, web, parser)
        {            
        }

        public override void AddPage(ListItem item, ProvisioningTemplate template)
        {
            var modelProvider = new FileModelProvider(this.Web, template.Connector);
            string pageUrl = GetUrl(item, false);
            Model.File file = modelProvider.GetFile(pageUrl, this.TokenParser);
            if (null == template.Files.Find((f) => string.Equals(f.Src, file.Src, System.StringComparison.OrdinalIgnoreCase)))
            {
                ObjectFiles.CreateLocalFile(this.Web, pageUrl, template.Connector);
                template.Files.Add(file);
            }
        }
    }
}
