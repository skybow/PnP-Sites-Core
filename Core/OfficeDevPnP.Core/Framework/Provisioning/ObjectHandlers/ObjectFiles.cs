using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Common;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.File;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.IO;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.Taxonomy;

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
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);
                List<string> filesList = new List<string>();

                foreach (var file in template.Files)
                {
                    try
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
                                SetFileProperties(targetFile, transformedProperties, false);
                            }

                            if (file.WebParts != null && file.WebParts.Any())
                            {
                                targetFile.EnsureProperties(f => f.ServerRelativeUrl);
                                filesList.Add(targetFile.ServerRelativeUrl);
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
                                        defaultWebPart.MoveWebPartTo(webpart.Zone, (int)webpart.Order);
                                        defaultWebPart.SaveWebPartChanges();
                                        SetProperties(webpart.Contents, defaultWebPart, scope);
                                    }
                                    else
                                    {
                                        AddWebPart(web, parser, webpart, targetFile);
                                    }
                                }
                                web.Context.ExecuteQueryRetry();
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
                    catch (System.Exception ex)
                    {
                        
                    }
                }

                FixViewsAfterAddingWebParts(web, filesList, parser, template);
            }
            return parser;
        }

        protected void FixViewsAfterAddingWebParts(Web web, List<string> files, TokenParser parser, ProvisioningTemplate template)
        {
            web.Context.Load(web.Lists, w => w.IncludeWithDefaultProperties(l => l.Views));
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQuery();
            bool isDirty = false;

            foreach (var list in web.Lists)
            {
                var views = list.Views;
                bool viewUpdated = false;
                foreach (var view in views)
                {
                    var url = view.ServerRelativeUrl;
                    bool exist = files.Any(f => f.Equals(url, System.StringComparison.InvariantCultureIgnoreCase));
                    var listFromTemplate = template.Lists.FirstOrDefault(l => l.Title == list.Title);
                    Model.View viewFromTemplate = listFromTemplate == null || listFromTemplate.Views == null 
                        ? null 
                        : listFromTemplate.Views.FirstOrDefault(v => 
                            v.PageUrl != null && v.PageUrl.EndsWith(url.Substring(web.ServerRelativeUrl.Length))
                            );
                    if (exist && viewFromTemplate != null && string.IsNullOrEmpty(view.Title)) { 
                        var viewElement = XElement.Parse(viewFromTemplate.SchemaXml);
                        var displayNameElement = viewElement.Attribute("DisplayName");
                        view.Title = displayNameElement == null ? Path.GetFileNameWithoutExtension(viewFromTemplate.PageUrl) : displayNameElement.Value;
                        view.Hidden = false;
                        view.Update();
                        isDirty = true;
                        viewUpdated = true;
                    }
                }
                if (viewUpdated)
                {
                    list.Update();
                }
            }

            if (isDirty) {
                web.Context.ExecuteQuery();
            }

        }

        public void SetFileProperties(File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            var context = file.Context;
            if (properties != null && properties.Count > 0)
            {
                // Get a reference to the target list, if any
                // and load file item properties
                var parentList = file.ListItemAllFields.ParentList;
                context.Load(parentList);
                context.Load(file.ListItemAllFields);
                try
                {
                    context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    // If this throws ServerException (does not belong to list), then shouldn't be trying to set properties)
                    if (ex.Message != "The object specified does not belong to a list.")
                    {
                        throw;
                    }
                }

                // Loop through and detect changes first, then, check out if required and apply
                foreach (var kvp in properties)
                {
                    var propertyName = kvp.Key;
                    var propertyValue = kvp.Value;

                    var fieldValues = file.ListItemAllFields.FieldValues;
                    var targetField = parentList.Fields.GetByInternalNameOrTitle(propertyName);
                    targetField.EnsureProperties(f => f.TypeAsString, f => f.ReadOnlyField);

                    if (true)  // !targetField.ReadOnlyField)
                    {
                        switch (propertyName.ToUpperInvariant())
                        {
                            case "CONTENTTYPE":
                                {
                                    Microsoft.SharePoint.Client.ContentType targetCT = parentList.GetContentTypeByName(propertyValue);
                                    context.ExecuteQueryRetry();

                                    if (targetCT != null)
                                    {
                                        file.ListItemAllFields["ContentTypeId"] = targetCT.StringId;
                                    }
                                    else
                                    {
                                        Log.Error(Constants.LOGGING_SOURCE, "Content Type {0} does not exist in target list!", propertyValue);
                                    }
                                    break;
                                }
                            default:
                                {
                                    switch (targetField.TypeAsString)
                                    {
                                        case "User":
                                            var user = parentList.ParentWeb.EnsureUser(propertyValue);
                                            context.Load(user);
                                            context.ExecuteQueryRetry();

                                            if (user != null)
                                            {
                                                var userValue = new FieldUserValue
                                                {
                                                    LookupId = user.Id,
                                                };
                                                file.ListItemAllFields[propertyName] = userValue;
                                            }
                                            break;
                                        case "URL":
                                            var urlArray = propertyValue.Split(',');
                                            var linkValue = new FieldUrlValue();
                                            if (urlArray.Length == 2)
                                            {
                                                linkValue.Url = urlArray[0];
                                                linkValue.Description = urlArray[1];
                                            }
                                            else
                                            {
                                                linkValue.Url = urlArray[0];
                                                linkValue.Description = urlArray[0];
                                            }
                                            file.ListItemAllFields[propertyName] = linkValue;
                                            break;
                                        case "LookupMulti":
                                            var lookupMultiValue = JsonUtility.Deserialize<FieldLookupValue[]>(propertyValue);
                                            file.ListItemAllFields[propertyName] = lookupMultiValue;
                                            break;
                                        case "TaxonomyFieldType":
                                            var taxonomyValue = JsonUtility.Deserialize<TaxonomyFieldValue>(propertyValue);
                                            file.ListItemAllFields[propertyName] = taxonomyValue;
                                            break;
                                        case "TaxonomyFieldTypeMulti":
                                            var taxonomyValueArray = JsonUtility.Deserialize<TaxonomyFieldValue[]>(propertyValue);
                                            file.ListItemAllFields[propertyName] = taxonomyValueArray;
                                            break;
                                        default:
                                            file.ListItemAllFields[propertyName] = propertyValue;
                                            break;
                                    }
                                    break;
                                }
                        }
                    }
                    file.ListItemAllFields.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }

        private static void SetProperties(string xml, WebPartDefinition defaultWebPart, PnPMonitoredScope scope)
        {
            try
            {
                var defaultProperties = defaultWebPart.WebPart.Properties;
                if (!xml.Contains("http://schemas.microsoft.com/WebPart/v3")) return;

                using (var reader = new StringReader(xml))
                {
                    XElement xelement = XElement.Load(reader);
                    var webPart = xelement.Elements().First();
                    var data = webPart.Elements().FirstOrDefault(x => x.Name.LocalName == "data");
                    var propertiesElement = data.Elements().FirstOrDefault(x => x.Name.LocalName == "properties");
                    IEnumerable<XElement> properties = propertiesElement.Elements();
                    foreach (var property in properties)
                    {
                        var propertyName = property.Attribute("name").Value;
                        var propertyValue = property.Value;
                        if (defaultProperties.FieldValues.ContainsKey(propertyName))
                        {
                            defaultProperties[propertyName] = Convert.ChangeType(propertyValue, Type.GetTypeCode(defaultProperties[propertyName].GetType()));
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                scope.LogError(exception, "Resetting properties for default Web part error");
            }

        }

        private static bool GetIsDefaultWebPart(Microsoft.SharePoint.Client.WebParts.WebPart webPart)
        {
            if (!webPart.Properties.FieldValues.ContainsKey("Default")) return false;
            var isDefaultString = webPart.Properties["Default"].ToString();
            return !string.IsNullOrEmpty(isDefaultString) && bool.Parse(isDefaultString);
        }

        private static void AddWebPart(Web web, TokenParser parser, OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart webPart, File targetFile)
        {
            var webPartPage = web.GetFileByServerRelativeUrl(targetFile.ServerRelativeUrl);

            var xml = parser.ParseString(webPart.Contents, Guid.Empty.ToString());
            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(xml);

            limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, webPart.Zone, (int)webPart.Order);
        }

        private static bool CheckOutIfNeeded(Web web, File targetFile)
        {
            var checkedOut = false;

            try
            {
                web.Context.Load(targetFile, f => f.CheckOutType, f => f.ListItemAllFields, f => f.ListItemAllFields.ParentList.ForceCheckout);
                web.Context.ExecuteQueryRetry();
                if (targetFile.ListItemAllFields.ServerObjectIsNull.HasValue
                    && !targetFile.ListItemAllFields.ServerObjectIsNull.Value
                    && targetFile.ListItemAllFields.ParentList.ForceCheckout)
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
                var files = new OfficeDevPnP.Core.Framework.Provisioning.Model.FileCollection(template);
                var modelProvider = new FileModelProvider(web, connector);
                var parser = new TokenParser(web, new ProvisioningTemplate());
                foreach (var provider in providers)
                {
                    var pageUrl = provider.GetUrl();
                    try
                    {
                        var file = modelProvider.GetFile(pageUrl, parser);
                        if (null != file)
                        {
                            files.Add(file);
                            this.CreateLocalFile(web, pageUrl, connector);
                        }
                        else
                        {
                            scope.LogError("File does not exist. URL:{0}.", pageUrl);
                        }
                    }
                    catch (Exception exception)
                    {
                        scope.LogError(exception, "Export file error: {0}", pageUrl);
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

        private System.Collections.Generic.List<IUrlProvider> GetUrlProviders(ListInstanceCollection templateLists)
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
