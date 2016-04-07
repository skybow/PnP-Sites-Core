using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.Page;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using File = Microsoft.SharePoint.Client.File;
using WebPart = OfficeDevPnP.Core.Framework.Provisioning.Model.WebPart;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Pages"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.RootFolder.WelcomePage);

                foreach (var page in template.Pages)
                {
                    var url = parser.ParseString(page.Url);

                    if (!url.ToLower().StartsWith(web.ServerRelativeUrl.ToLower()))
                    {
                        url = UrlUtility.Combine(web.ServerRelativeUrl, url);
                    }

                    var exists = true;
                    File file = null;
                    try
                    {
                        file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (ServerException ex)
                    {
                        if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                        {
                            exists = false;
                        }
                    }
                    if (exists)
                    {
                        if (page.Overwrite)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0_, url);

                                string welcomePageUrl = string.IsNullOrEmpty(web.RootFolder.WelcomePage) ? "" : UrlUtility.Combine(web.ServerRelativeUrl, web.RootFolder.WelcomePage);
                                if (!string.IsNullOrEmpty(welcomePageUrl) && url.Equals(welcomePageUrl, StringComparison.InvariantCultureIgnoreCase))
                                    web.SetHomePage(string.Empty);

                                file.DeleteObject();
                                web.Context.ExecuteQueryRetry();
                                this.AddPage(web, url, page, parser);
                                }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0_, url);
                            this.AddPage(web, url, page, parser);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                        }
                    }

                    if (page.WelcomePage)
                    {
                        var rootFolderRelativeUrl = url.Substring(web.ServerRelativeUrl.Length + 1);
                        web.SetHomePage(rootFolderRelativeUrl);
                    }

                    if (page.WebParts != null & page.WebParts.Any())
                    {
                        this.AddWebParts(web, page, parser);
                    }
                    if (page.Security != null)
                    {
                        file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file.ListItemAllFields);
                        web.Context.ExecuteQueryRetry();
                        file.ListItemAllFields.SetSecurity(parser, page.Security);
                    }
                }
            }
            return parser;
        }

        //TODO: move to class
        private void AddPage(Web web, string url, Page page, TokenParser parser)
        {
            var publishingPage = page as PublishingPage;
            if (publishingPage != null)
            {
                string layoutUrl = parser.ParseString(publishingPage.PageLayoutUrl);
                web.AddPublishingPageByUrl(url, layoutUrl, publishingPage.PageTitle, publishingPage.Html);
            }
            else
            {
                var contentPage = page as ContentPage;
                if (contentPage != null)
                {
                    web.AddWikiPageByUrl(url, contentPage.Html);
                }
                else
                {
                    web.AddWikiPageByUrl(url);
                    web.AddLayoutToWikiPage(page.Layout, url);
                }
            }
        }

        //TODO: refactor this
        private void AddWebParts(Web web, Page page, TokenParser parser)
        {
            var url = parser.ParseString(page.Url);
            ContentPage contentPage = page as ContentPage;
            if (contentPage != null)
            {
                var file = web.GetFileByServerRelativeUrl(url);
                file.CheckOut();
                foreach (var model in contentPage.WebParts)
                {
                    try
                    {
                        model.Contents = parser.ParseString(model.Contents);

                        string oldId = null;
                        string newId = null;
                        if (!WebPartsModelProvider.IsV3FormatXml(model.Contents))
                    {
                            var id = WebPartsModelProvider.GetWebPartControlId(model.Contents);
                            var idToReplace = GetNewControlId();
                            model.Contents = model.Contents.Replace(id, idToReplace);
                            newId = this.GetIdFromControlId(idToReplace);
                            oldId = this.GetIdFromControlId(id);
                        }

                        var addedWebPart = this.AddWebPart(web, model, file);
                        newId = newId ?? addedWebPart.Id.ToString().ToLower();
                        oldId = oldId ?? GetWebPartIdFromSchema(model.Contents).ToLower();
                        parser.AddToken(new IdToken(web, newId, oldId));
                    }
                    catch (Exception ex)
                        {
                        Log.Error(Constants.LOGGING_SOURCE_FRAMEWORK_PROVISIONING, "Could not add webpart: {0} - {1}", ex.Message, ex.StackTrace);
                        }
                    }

                var html = parser.ParseString(contentPage.Html);
                if (!string.IsNullOrEmpty(html))
                    {
                    web.AddHtmlToWikiPage(url, html);
                }
                file.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                return;
            }
                    }
        private string GetNewControlId()
        {
            return string.Format("g_{0}", Guid.NewGuid().ToString("D").Replace("-", "_"));
                }
        private string GetIdFromControlId(string controlId)
        {
            return controlId.Replace("g_", string.Empty).Replace("_", "-");
            }

        private WebPartDefinition AddWebPart(Web web, WebPart webPart, File pageFile)
        {
            LimitedWebPartManager limitedWebPartManager = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPart.Contents);
            WebPartDefinition wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, webPart.Zone, (int)webPart.Order);
            web.Context.Load(wpdNew, x => x.Id);
            web.Context.ExecuteQueryRetry();
            return wpdNew;
        }

        private string GetWebPartIdFromSchema(string xml)
        {
            return new Regex(@"(?<=webpartid="").*?(?="")", RegexOptions.Singleline).Match(xml).Value;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var lists = this.GetListsWithPages(template);
                template.Pages = new PageCollection(template);
                var parser = new TokenParser(web, new ProvisioningTemplate());

                var homePageUrl = web.GetHomePageRelativeUrl();
                foreach (var list in lists)
                {
                    try
                    {
                        List splist = web.Lists.GetById(list.ID);
                        web.Context.Load(splist);
                        web.Context.ExecuteQueryRetry();
                        if (!creationInfo.ExecutePreProvisionEvent<ListInstance, List>(Handlers.Pages, template, list, null))
                        {
                            continue;
                        }

                        var listItems = GetListPages(web, splist);
                        var fileItems = listItems.AsEnumerable().Where(x => x.IsFile());
                        foreach (ListItem item in fileItems)
                        {
                            try
                            {
                                IPageModelProvider provider = GetProvider(item, homePageUrl, web, parser);
                                if (null != provider)
                                {
                                    provider.AddPage(item, template);
                                }
                            }
                            catch (Exception ex)
                            {
                                var message = string.Format("Error in export page for list: {0}", list.ServerRelativeUrl);
                                scope.LogError(ex, message);
                            }
                        }

                        creationInfo.ExecutePostProvisionEvent<ListInstance, List>(Handlers.Pages, template, list, splist);
                    }
                    catch (Exception exception)
                    {
                        var message = string.Format("Error in export publishing page for list: {0}", list.ServerRelativeUrl);
                        scope.LogError(exception, message);
                    }
                }                
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        internal static IPageModelProvider GetProvider(ListItem item, string homePageUrl, Web web, TokenParser parser)
        {
            var fieldValues = item.FieldValues;
            IPageModelProvider provider = null;

            if (fieldValues.ContainsKey("PublishingPageContent") && 
                fieldValues.ContainsKey("PublishingPageLayout") && (null != fieldValues["PublishingPageLayout"]))
            {
                provider = new PublishingPageModelProvider(homePageUrl, web, parser);
            }
            else if (fieldValues.ContainsKey("WikiField"))
            {
                provider = new ContentPageModelProvider(homePageUrl, web, parser);
            }
            else
            {
                provider = new WebPartPageModelProvider(homePageUrl, web, parser);
            }

            return provider;
        }

        private IEnumerable<ListItem> GetListPages(Web web, List list)
        {
            var caml = CamlQuery.CreateAllItemsQuery();
            var listItems = list.GetItems(caml);

            web.Context.Load(listItems, includes => includes.Include(i => i.File.Versions));
            web.Context.Load(listItems);
            web.Context.ExecuteQueryRetry();

            var fileItems = listItems.AsEnumerable().Where(x => x.IsFile());
            return fileItems;
        }


        private IEnumerable<ListInstance> GetListsWithPages(ProvisioningTemplate template)
        {
            return template.Lists.Where(x => x.TemplateType == (int)ListTemplateType.WebPageLibrary
                || x.TemplateType == (int)ListTemplateType.HomePageLibrary
#if CLIENTSDKV15
 || x.TemplateType == (int)ListTemplateType.PublishingPages
#endif
);
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Pages.Any();
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
