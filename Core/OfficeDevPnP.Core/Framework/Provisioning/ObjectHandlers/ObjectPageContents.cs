using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client.WebParts;
using System.Xml.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.IO;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPageContents : ObjectContentHandlerBase
    {
        public override string Name
        {
            get { return "Page Contents"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var homepageUrl = web.RootFolder.WelcomePage;
                if (string.IsNullOrEmpty(homepageUrl))
                {
                    homepageUrl = "Default.aspx";
                }
                var welcomePageUrl = UrlUtility.Combine(web.ServerRelativeUrl, homepageUrl);

                bool welcomePageAdded = false;

                var file = web.GetFileByServerRelativeUrl(welcomePageUrl);
                try
                {
                    var listItem = file.EnsureProperty(f => f.ListItemAllFields);
                    if (listItem != null)
                    {
                        TokenParser parser = new TokenParser(web, new ProvisioningTemplate());
                        Export.Page.IPageModelProvider provider = ObjectPages.GetProvider(listItem, homepageUrl, web, parser);
                        if (null != provider)
                        {
                            provider.AddPage(listItem, template);
                            welcomePageAdded = true;
                        }
                    }
                }
                catch (ServerException ex)
                {
                    if (ex.ServerErrorCode != -2146232832)
                    {
                        throw;
                    }
                    else
                    {
                        // Page does not belong to a list, extract the file as is
                        template = GetFileContents(web, template, welcomePageUrl, creationInfo, scope);
                        welcomePageAdded = true;
                    }
                }

                if (welcomePageAdded)
                {
                    if (template.WebSettings == null)
                    {
                        template.WebSettings = new WebSettings();
                    }
                    template.WebSettings.WelcomePage = homepageUrl;
                }

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }
        
        private ProvisioningTemplate GetFileContents(Web web, ProvisioningTemplate template, string welcomePageUrl, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var fullUri = new Uri(UrlUtility.Combine(web.Url, web.RootFolder.WelcomePage));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

            var webParts = web.GetWebParts(welcomePageUrl);

            var file = web.GetFileByServerRelativeUrl(welcomePageUrl);

            var homeFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = fileName,
                Overwrite = true,
            };            
            // Add field values to file
            //RetrieveFieldValues(web, file, homeFile);

            WebPartsModelProvider webpartsProvider = new WebPartsModelProvider(web);
            var webparts = webpartsProvider.Retrieve(welcomePageUrl, new TokenParser(web, new ProvisioningTemplate()));
            homeFile.WebParts.AddRange(webparts);

            template.Files.Add(homeFile);

            // Persist file using connector
            if (creationInfo.PersistBrandingFiles)
            {
                PersistFile(web, creationInfo, scope, folderPath, fileName);
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = web.Context.Credentials != null ? true : false;
            }
            return _willExtract.Value;
        }

    }
}
