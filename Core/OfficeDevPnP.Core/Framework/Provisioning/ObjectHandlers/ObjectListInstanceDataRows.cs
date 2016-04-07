using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.ListContent;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectListInstanceDataRows : ObjectHandlerBase
    {
        private Dictionary<Guid, ListItemsProvider> m_listContentProviders = null;

        public override string Name
        {
            get { return "List instances Data Rows"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Lists.Any())
                {
                    var rootWeb = (web.Context as ClientContext).Site.RootWeb;

                    web.EnsureProperties(w => w.ServerRelativeUrl);
                    
                    web.Context.Load(web.Lists, lc => lc.IncludeWithDefaultProperties(l => l.RootFolder.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    var existingLists = web.Lists.AsEnumerable<List>().Select(existingList => existingList.RootFolder.ServerRelativeUrl).ToList();
                    var serverRelativeUrl = web.ServerRelativeUrl;

                    #region DataRows

                    foreach (var listInstance in template.Lists)
                    {
                        if (listInstance.DataRows != null && listInstance.DataRows.Any())
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Processing_data_rows_for__0_, listInstance.Title);
                            // Retrieve the target list
                            var list = web.Lists.GetByTitle(listInstance.Title);
                            web.Context.Load(list);
                            web.Context.ExecuteQueryRetry();
                            
                            ListItemsProvider provider = new ListItemsProvider(list, web, template);                            
                            provider.AddListItems(listInstance.DataRows, template, parser, scope);
                            if (null == m_listContentProviders)
                            {
                                m_listContentProviders = new Dictionary<Guid, ListItemsProvider>();
                            }
                            m_listContentProviders[list.Id] = provider;
                        }
                    }

                    UpdateLookupValues(web, scope);

                    #endregion
                }
            }

            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                foreach (var listInstance in template.Lists)
                {
                    List list = web.Lists.GetById(listInstance.ID);
                    web.Context.Load(list);
                    web.Context.ExecuteQueryRetry();

                    if (creationInfo.ExecutePreProvisionEvent<ListInstance, List>(Handlers.ListContents, template, listInstance, list))                    
                    {
                        ListItemsProvider provider = new ListItemsProvider(list, web, template);
                        List<DataRow> dataRows = provider.ExtractItems(creationInfo, scope);
                        listInstance.DataRows.AddRange(dataRows);

                        creationInfo.ExecutePostProvisionEvent<ListInstance, List>(Handlers.ListContents, template, listInstance, list);
                    }
                }
            }
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Lists.Any(l => l.DataRows.Any());
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

        private void UpdateLookupValues(Web web, PnPMonitoredScope scope)
        {
            if (null != m_listContentProviders)
            {
                foreach (KeyValuePair<Guid, ListItemsProvider> pair in m_listContentProviders)
                {
                    Guid listId = pair.Key;
                    ListItemsProvider provider = pair.Value;

                    provider.UpdateLookups(GetLookupDependentProvider, scope);
                }                
            }
        }

        private ListItemsProvider GetLookupDependentProvider(Guid listId)
        {
            ListItemsProvider provider;
            if( (null != m_listContentProviders)&& m_listContentProviders.TryGetValue(listId, out provider ) )
            {
                return provider;
            }
            return null;
        }
    }
}

