using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Common;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.File;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using WebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectFiles : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Files"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.ServerRelativeUrl);

                foreach (var file in template.Files)
                {
                    var folderName = parser.ParseString(file.Folder);

                    if (folderName.ToLower().StartsWith(web.ServerRelativeUrl.ToLower()))
                    {
                        folderName = folderName.Substring(web.ServerRelativeUrl.Length);
                    }

                    var folder = web.EnsureFolderPath(folderName);

                    File targetFile = null;

                    var checkedOut = false;

                    targetFile = folder.GetFile(template.Connector.GetFilenamePart(file.Src));

                    if (targetFile != null)
                    {
                        if (file.Overwrite)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_and_overwriting_existing_file__0_, file.Src);
                            checkedOut = CheckOutIfNeeded(web, targetFile);

                            using (var stream = template.Connector.GetFileStream(file.Src))
                            {
                                targetFile = folder.UploadFile(template.Connector.GetFilenamePart(file.Src), stream, file.Overwrite);
                            }
                        }
                        else
                        {
                            checkedOut = CheckOutIfNeeded(web, targetFile);
                        }
                    }
                    else
                    {
                        using (var stream = template.Connector.GetFileStream(file.Src))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_file__0_, file.Src);
                            targetFile = folder.UploadFile(template.Connector.GetFilenamePart(file.Src), stream, file.Overwrite);
                        }

                        checkedOut = CheckOutIfNeeded(web, targetFile);
                    }

                    if (targetFile != null)
                    {
                        if (file.Properties != null && file.Properties.Any())
                        {
                            Dictionary<string, string> transformedProperties = file.Properties.ToDictionary(property => property.Key, property => parser.ParseString(property.Value));
                            targetFile.SetFileProperties(transformedProperties, false); // if needed, the file is already checked out
                        }

                        if (file.WebParts != null && file.WebParts.Any())
                        {
                            targetFile.EnsureProperties(f => f.ServerRelativeUrl);

                            var existingWebParts = web.GetWebParts(targetFile.ServerRelativeUrl);

                            var enumerator = existingWebParts.GetEnumerator();
                            var needToExecute = false;
                            WebPartDefinition defaultWebPart = null;
                            while (enumerator.MoveNext())
                            {
                                var current = enumerator.Current;
                                if (GetIsDefaultWebPart(current.WebPart))
                                {
                                    defaultWebPart = current;
                                }
                                else
                                {
                                    enumerator.Current.DeleteWebPart();
                                    needToExecute = true;
                                }
                            }

                            if (needToExecute)
                            {
                                web.Context.ExecuteQueryRetry();
                            }

                            foreach (var webpart in file.WebParts)
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Adding_webpart___0___to_page, webpart.Title);

                                if (defaultWebPart != null && defaultWebPart.WebPart.Title == webpart.Title)
                                {
                                    defaultWebPart.MoveWebPartTo(webpart.Zone, (int) webpart.Order);
                                    defaultWebPart.SaveWebPartChanges();
                                }
                                else
                                {
                                    AddWebPart(web, parser, webpart, targetFile);
                                }
                            }
                            web.Context.ExecuteQuery();
                        }

                        if (checkedOut)
                        {
                            targetFile.CheckIn("", CheckinType.MajorCheckIn);
                            web.Context.ExecuteQueryRetry();
                        }

                        // Don't set security when nothing is defined. This otherwise breaks on files set outside of a list
                        if (file.Security != null &&
                            (file.Security.ClearSubscopes == true || file.Security.CopyRoleAssignments == true || file.Security.RoleAssignments.Count > 0))
                        {
                            targetFile.ListItemAllFields.SetSecurity(parser, file.Security);
                        }
                    }
                }
            }
            return parser;
        }

        private static bool GetIsDefaultWebPart(Microsoft.SharePoint.Client.WebParts.WebPart webPart)
        {
            if (!webPart.Properties.FieldValues.ContainsKey("Default")) return false;
            var isDefaultString = webPart.Properties["Default"].ToString();
            return !string.IsNullOrEmpty(isDefaultString) && bool.Parse(isDefaultString);
        }

        private static void AddWebPart(Web web, TokenParser parser, WebPart webPart, File targetFile)
        {
            var webPartPage = web.GetFileByServerRelativeUrl(targetFile.ServerRelativeUrl);

            var xml = parser.ParseString(webPart.Contents, Guid.Empty.ToString());
            LimitedWebPartManager  limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(xml);

            limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, webPart.Zone, (int) webPart.Order);
        }

        private static bool CheckOutIfNeeded(Web web, File targetFile)
        {
            var checkedOut = false;

            try
            {
                web.Context.Load(targetFile, f => f.CheckOutType, f => f.ListItemAllFields, f => f.ListItemAllFields.ParentList.ForceCheckout);
                web.Context.ExecuteQueryRetry();
                if (IsLibraryFile(targetFile))
                {
                    if (targetFile.CheckOutType == CheckOutType.None)
                    {
                        targetFile.CheckOut();
                    }
                    checkedOut = true;
                }
            }
            catch (ServerException ex)
            {
                // Handling the exception stating the "The object specified does not belong to a list."
                if (ex.ServerErrorCode != -2146232832)
                {
                    throw;
                }
            }
            return checkedOut;
        }

        private static bool IsLibraryFile(File file)
        {
            return file.ListItemAllFields.FieldValues.Any();
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            web.EnsureProperties(x => x.Url, x => x.Id);
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var connector = template.Connector;
                var providers = this.GetUrlProviders(template.Lists);
                var files = new List<Model.File>();
                var modelProvider = new FileModelProvider(web, connector);
                foreach (var provider in providers)
                {
                    try
                    {
                        var pageUrl = provider.GetUrl();
                        var file = modelProvider.GetFile(provider.GetUrl());
                        files.Add(file);

                        this.CreateLocalFile(web, pageUrl, connector);
                    }
                    catch (Exception exception)
                    {
                        scope.LogError(exception, "Export file error");
                    }
                }

                template.Files = files;

                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private List<IUrlProvider> GetUrlProviders(List<ListInstance> templateLists)
        {
            var result = new List<IUrlProvider>();
            var forms = templateLists.SelectMany(x => x.Forms);
            var views = templateLists.SelectMany(x => x.Views);
            result.AddRange(forms);
            result.AddRange(views);
            return result;
        }

        private void CreateLocalFile(Web web, string pageUrl, FileConnectorBase connector)
        {
            var fileContent = web.GetPageContentXmlWithoutWebParts(pageUrl);
            var fileName = Path.GetFileName(pageUrl);
            var folderPath = Path.GetDirectoryName(pageUrl).TrimStart('\\');

            Byte[] info = new UTF8Encoding(true).GetBytes(fileContent);
            using (var stream = new MemoryStream(info))
            {
                connector.SaveFileStream(fileName, folderPath, stream);
            }
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Files.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var collList = web.Lists;
                var lists = web.Context.LoadQuery(collList);

                web.Context.ExecuteQueryRetry();

                _willExtract = lists.Any();
            }
            return _willExtract.Value;
        }
    }
}
