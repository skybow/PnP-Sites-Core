﻿using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.IO;
using OfficeDevPnP.Core.Diagnostics;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectComposedLook : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Composed Looks"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.ComposedLook != null &&
                    !template.ComposedLook.Equals(ComposedLook.Empty))
                {
                    try
                    {

                    bool executeQueryNeeded = false;
                    if (executeQueryNeeded)
                    {
                        web.Context.ExecuteQueryRetry();
                    }

                    if (String.IsNullOrEmpty(template.ComposedLook.ColorFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.FontFile) &&
                        String.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                    {
                        // Apply OOB theme
                        web.SetComposedLookByUrl(template.ComposedLook.Name);
                    }
                    else
                    {
                        // Apply custom theme
                        string colorFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            colorFile = parser.ParseString(template.ComposedLook.ColorFile);
                        }
                        string backgroundFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            backgroundFile = parser.ParseString(template.ComposedLook.BackgroundFile);
                        }
                        string fontFile = null;
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            fontFile = parser.ParseString(template.ComposedLook.FontFile);
                        }

                        string masterUrl = null;
                        if (template.WebSettings != null && !string.IsNullOrEmpty(template.WebSettings.MasterPageUrl))
                        {
                            masterUrl = parser.ParseString(template.WebSettings.MasterPageUrl);
                        }
                        web.CreateComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);
                        web.SetComposedLookByUrl(template.ComposedLook.Name, colorFile, fontFile, backgroundFile, masterUrl);

                        var composedLookJson = JsonConvert.SerializeObject(template.ComposedLook);

                        web.SetPropertyBagValue("_PnP_ProvisioningTemplateComposedLookInfo", composedLookJson);
                    }

                    // Persist composed look info in property bag

                }
                    catch (Exception ex)
                    {
                        Log.Error(Constants.LOGGING_SOURCE_FRAMEWORK_PROVISIONING, "Could not provision ComposedLook: {0} - {1}", ex.Message, ex.StackTrace);
            }

                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Retrieving_current_composed_look);

                // Ensure that we have URL property loaded for web and site
                web.EnsureProperty(w => w.Url);
                Site site = (web.Context as ClientContext).Site;
                site.EnsureProperty(s => s.Url);
               
                SharePointConnector spConnector = new SharePointConnector(web.Context, web.Url, "dummy");
                // to get files from theme catalog we need a connector linked to the root site
                SharePointConnector spConnectorRoot;
                if (!site.Url.Equals(web.Url, StringComparison.InvariantCultureIgnoreCase))
                {
                    spConnectorRoot = new SharePointConnector(web.Context.Clone(site.Url), site.Url, "dummy");
                }
                else
                {
                    spConnectorRoot = spConnector;
                }

                // Check if we have composed look info in the property bag, if so, use that, otherwise try to detect the current composed look
                if (web.PropertyBagContainsKey("_PnP_ProvisioningTemplateComposedLookInfo"))
                {
                    scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Using_ComposedLookInfoFromPropertyBag);

                    try
                    {
                        var composedLook = JsonConvert.DeserializeObject<ComposedLook>(web.GetPropertyBagValueString("_PnP_ProvisioningTemplateComposedLookInfo", ""));
                        if (composedLook.Name == null || composedLook.BackgroundFile == null || composedLook.FontFile == null)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_ComposedLookInfoFailedToDeserialize);
                            throw new JsonSerializationException();
                        }

                        composedLook.BackgroundFile = TokenizeUrl(composedLook.BackgroundFile, parser);
                        composedLook.FontFile = TokenizeUrl(composedLook.FontFile, parser);
                        composedLook.ColorFile = TokenizeUrl(composedLook.ColorFile, parser);
                        template.ComposedLook = composedLook;

                        if (!web.IsSubSite() && creationInfo != null && 
                                creationInfo.PersistBrandingFiles && creationInfo.FileConnector != null)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                            // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, parser.ParseString(composedLook.BackgroundFile), scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, parser.ParseString(composedLook.ColorFile), scope);
                            DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, parser.ParseString(composedLook.FontFile), scope);
                        }
                        // Create file entries for the custom theme files  
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.BackgroundFile, template.Connector);
                            f.Folder = TokenizeUrl(f.Folder, parser);
                            template.Files.Add(f);
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.ColorFile, template.Connector);
                            f.Folder = TokenizeUrl(f.Folder, parser);
                            template.Files.Add(f);
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            var f = GetComposedLookFile(template.ComposedLook.FontFile, template.Connector);
                            f.Folder = TokenizeUrl(f.Folder, parser);
                            template.Files.Add(f);
                        }

                    }
                    catch (JsonSerializationException)
                    {
                        // cannot deserialize the object, fall back to composed look detection
                        template = DetectComposedLook(web, template, parser, creationInfo, scope, spConnector, spConnectorRoot);
                    }

                }
                else
                {
                    template = DetectComposedLook(web, template, parser, creationInfo, scope, spConnector, spConnectorRoot);
                }

                if (creationInfo != null && creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate DetectComposedLook(Web web, ProvisioningTemplate template, TokenParser parser,
                                                    ProvisioningTemplateCreationInformation creationInfo, 
                                                    PnPMonitoredScope scope, SharePointConnector spConnector, 
                                                    SharePointConnector spConnectorRoot)
        {

            var theme = web.GetCurrentComposedLook();

            if (theme != null)
            {
                if (creationInfo != null)
                {
                    // Don't exclude the DesignPreviewThemedCssFolderUrl property bag, if any
                    creationInfo.PropertyBagPropertiesToPreserve.Add("DesignPreviewThemedCssFolderUrl");
                }

                template.ComposedLook.Name = theme.Name;

                if (theme.IsCustomComposedLook)
                {
                    // Set the URL pointers to files
                    template.ComposedLook.BackgroundFile = FixFileUrl(TokenizeUrl(theme.BackgroundImage, parser));
                    template.ComposedLook.ColorFile = FixFileUrl(TokenizeUrl(theme.Theme, parser));
                    template.ComposedLook.FontFile = FixFileUrl(TokenizeUrl(theme.Font, parser));

                    // Download files if this is root site, since theme files are only stored there
                    if (!web.IsSubSite() && creationInfo != null && 
                        creationInfo.PersistBrandingFiles && creationInfo.FileConnector != null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_ExtractObjects_Creating_SharePointConnector);
                        // Let's create a SharePoint connector since our files anyhow are in SharePoint at this moment
                        // Download the theme/branding specific files
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, theme.BackgroundImage, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, theme.Theme, scope);
                        DownLoadFile(spConnector, spConnectorRoot, creationInfo.FileConnector, web, theme.Font, scope);
                    }

                    // Create file entries for the custom theme files, but only if it's a root site
                    // If it's root site we do not extract or set theme files, since those are in the root of the site collection
                    if (!web.IsSubSite())
                    {   
                        if (!string.IsNullOrEmpty(template.ComposedLook.BackgroundFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.BackgroundFile, template.Connector));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.ColorFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.ColorFile, template.Connector));
                        }
                        if (!string.IsNullOrEmpty(template.ComposedLook.FontFile))
                        {
                            template.Files.Add(GetComposedLookFile(template.ComposedLook.FontFile, template.Connector));
                        }
                    }
                    // If a base template is specified then use that one to "cleanup" the generated template model
                    if (creationInfo != null && creationInfo.BaseTemplate != null)
                    {
                        template = CleanupEntities(template, creationInfo.BaseTemplate);
                    }
                }
                else
                {
                    template.ComposedLook.BackgroundFile = "";
                    template.ComposedLook.ColorFile = "";
                    template.ComposedLook.FontFile = "";
                }
            }
            else
            {
                template.ComposedLook = null;
            }

            return template;
        }

        private void DownLoadFile(SharePointConnector reader, SharePointConnector readerRoot, FileConnectorBase writer, Web web, string asset, PnPMonitoredScope scope)
        {

            // No file passed...leave
            if (String.IsNullOrEmpty(asset))
            {
                return;
            }

            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ComposedLooks_DownLoadFile_Downloading_asset___0_, asset);
            ;

            SharePointConnector readerToUse;
            Model.File f = GetComposedLookFile(asset, null);

            if (f.Folder.StartsWith(web.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                f.Folder = f.Folder.Substring(web.ServerRelativeUrl.Length);
            }

            // in case of a theme catalog we need to use the root site reader as that list only exists on root site level
            if (f.Folder.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                readerToUse = readerRoot;
            }
            else
            {
                readerToUse = reader;
            }

            using (Stream s = readerToUse.GetFileStream(f.Src, f.Folder))
            {
                if (s != null)
                {
                    writer.SaveFileStream(f.Src, s);
                }
            }
        }

        private String FixFileName(string originalFileName)
        {
            // if we've found the file use the provided writer to persist the downloaded file
            String regexStrip = @"(\\|/|:|\*|\?|""|>|<|\||=)*";
            String result = Regex.Replace(originalFileName.Substring(0,
                originalFileName.IndexOf("?") > 0 ? originalFileName.IndexOf("?") : originalFileName.Length),
                regexStrip, "", RegexOptions.IgnorePatternWhitespace);

            return (result);
        }
        private String FixFileUrl(string originalFileUrl)
        {
            if (string.IsNullOrEmpty(originalFileUrl))
            {
                return "";
            }

            String fileUrl = originalFileUrl.Substring(0, originalFileUrl.LastIndexOf("/"));
            String fileName = FixFileName(originalFileUrl.Substring(originalFileUrl.LastIndexOf("/") + 1));

            String result = String.Format("{0}/{1}", fileUrl, fileName);

            return (result);
        }

        private Model.File GetComposedLookFile(string asset, FileConnectorBase connector)
        {
            int index = asset.LastIndexOf("/");
            Model.File file = new Model.File();
            string fileSrc = FixFileName(asset.Substring(index + 1));            
            file.Src = (null != connector)?Path.Combine(connector.GetConnectionString(), fileSrc): fileSrc;
            file.Folder = asset.Substring(0, index);
            file.Overwrite = true;
            file.Security = null;
            return file;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            if (template.ComposedLook != null && baseTemplate.ComposedLook != null)
            {
                if (template.ComposedLook.Equals(baseTemplate.ComposedLook))
                {
                    template.ComposedLook = null;
                }

            }
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = (template.ComposedLook != null && !template.ComposedLook.Equals(ComposedLook.Empty));
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }
}
