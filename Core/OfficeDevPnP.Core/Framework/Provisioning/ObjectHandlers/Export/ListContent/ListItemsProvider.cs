using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Field = Microsoft.SharePoint.Client.Field;
using List = Microsoft.SharePoint.Client.List;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.ListContent
{
    public class ListItemsProvider
    {
        #region Internal Classes

        internal class LookupDataRef
        {
            public FieldLookup Field { get; private set; }
            public Dictionary<ListItem, object> ItemLookupValues { get; private set; }//<listitem,lookupvalue>

            public LookupDataRef(FieldLookup field)
            {
                this.Field = field;
                this.ItemLookupValues = new Dictionary<ListItem, object>();
            }
        }

        #endregion //Internal Classes

        #region Fields

        private Dictionary<string, FieldValueProvider> m_dictFieldValueProviders = null;

        private Dictionary<int, int> m_mappingIDs = null;
        private Dictionary<string, LookupDataRef> m_lookups = null;

        #endregion //Fields

        #region Properties

        public List List { get; private set; }
        public Web Web { get; private set; }

        public ClientRuntimeContext Context
        {
            get
            {
                return this.Web.Context;
            }
        }

        #endregion //Properties

        #region Constructors

        public ListItemsProvider(List list, Web web, ProvisioningTemplate template)
        {
            this.List = list;
            this.Web = web;
        }

        #endregion //Constructors

        #region Methods

        public void AddListItems(DataRowCollection dataRows, TokenParser parser, PnPMonitoredScope scope)
        {
            Microsoft.SharePoint.Client.FieldCollection fields = this.List.Fields;
            this.Context.Load(fields);
            this.Context.ExecuteQueryRetry();

            foreach (var dataRow in dataRows)
            {
                try
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_list_item__0_, dataRows.IndexOf(dataRow) + 1);
                    var listitemCI = new ListItemCreationInformation();
                    var listitem = this.List.AddItem(listitemCI);

                    foreach (var dataValue in dataRow.Values)
                    {
                        Field dataField = fields.FirstOrDefault(
                            f => f.InternalName == parser.ParseString(dataValue.Key));

                        if ((dataField != null) && CanFieldContentBeIncluded(dataField, false))
                        {
                            string fieldValue = parser.ParseString(dataValue.Value);
                            if (!string.IsNullOrEmpty(fieldValue))
                            {
                                FieldValueProvider valueProvider = GetFieldValueProvider(dataField, this.Web);
                                if (null != valueProvider)
                                {
                                    object itemValue = valueProvider.GetFieldValueTyped(fieldValue);
                                    if (null != itemValue)
                                    {
                                        if (dataField.FieldTypeKind == FieldType.Lookup)
                                        {
                                            FieldLookup lookupField = (FieldLookup)dataField;
                                            RegisterLookupReference(lookupField, listitem, itemValue);
                                        }
                                        else
                                        {
                                            listitem[dataField.InternalName] = itemValue;
                                        }                                        
                                    }                                    
                                }
                            }                                                     
                        }
                    }
                    listitem.Update();
                    this.Context.ExecuteQueryRetry(); // TODO: Run in batches?

                    AddIDMappingEntry(listitem, dataRow);

                    if (dataRow.Security != null && dataRow.Security.RoleAssignments.Count != 0)
                    {
                        listitem.SetSecurity(parser, dataRow.Security);
                    }
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ListInstancesDataRows_Creating_listitem_failed___0_____1_, ex.Message, ex.StackTrace);
                }
            }
        }

        public List<DataRow> ExtractItems()
        {
            List<DataRow> dataRows = new List<DataRow>();

            CamlQuery query = new CamlQuery()
            {
                DatesInUtc = true,
                ViewXml = "" //Should be recursive
            };
            ListItemCollection items = this.List.GetItems(query);            
            this.Context.Load(items, col => col.IncludeWithDefaultProperties(i => i.HasUniqueRoleAssignments));
            this.Context.ExecuteQueryRetry();

            List<Field> fields = GetListContentSerializableFields(true);
            foreach (ListItem item in items)
            {
                Dictionary<string, string> values = new Dictionary<string, string>();
                foreach (Field field in fields)
                {
                    if (CanFieldContentBeIncluded(field, true))
                    {
                        string str = "";
                        object value = null; ;
                        try
                        {
                            value = item[field.InternalName];
                        }
                        catch (Exception ex)
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, ex,
                                "Failed to read item field value. List:{0}, Item ID:{1}, Field: {2}", this.List.Title, item.Id, field.InternalName);
                        }
                        if (null != value)
                        {
                            try
                            {
                                FieldValueProvider provider = GetFieldValueProvider(field, this.Web);
                                str = provider.GetValidatedValue(value);
                            }
                            catch (Exception ex)
                            {
                                Log.Warning(Constants.LOGGING_SOURCE, ex,
                                    "Failed to serialize item field value. List:{0}, Item ID:{1}, Field: {2}", this.List.Title, item.Id, field.InternalName);
                            }
                            if (!string.IsNullOrEmpty(str))
                            {
                                values.Add(field.InternalName, str);
                            }
                            }
                        }
                    }

                if (values.Any())
                {
                    ObjectSecurity security = null;
                    if (item.HasUniqueRoleAssignments)
                    {
                        try
                        {
                            security = item.GetSecurity();
                            security.ClearSubscopes = true;
                }
                        catch (Exception ex)
                {
                            Log.Error(ex, Constants.LOGGING_SOURCE, "Failed to get item security. Item ID: {0}, List: '{1}'.", item.Id, this.List.Title);
                        }
                    }

                    DataRow row = new DataRow( values, security );
                    dataRows.Add(row);
                }
            }

            return dataRows;
        }

        public void UpdateLookups(Func<Guid, ListItemsProvider> fnGetLookupDependentProvider)
        {
            if (null != m_lookups)
            {
                bool valueUpdated = false;
                foreach (KeyValuePair<string, LookupDataRef> pair in m_lookups)
                {
                    LookupDataRef lookupData = pair.Value;
                    if (0 < lookupData.ItemLookupValues.Count)
                    {
                        Guid sourceListId = Guid.Empty;
                        try
                        {
                            sourceListId = new Guid(lookupData.Field.LookupList);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, Constants.LOGGING_SOURCE, "Failed to get source list for lookup field. Field Name: {0}, Source List: {1}.",
                                lookupData.Field.InternalName, lookupData.Field.LookupList);
                        }
                        if (!Guid.Empty.Equals(sourceListId))
                        {
                            ListItemsProvider sourceProvider = fnGetLookupDependentProvider(sourceListId);
                            if ((null != sourceProvider) && (null != sourceProvider.m_mappingIDs))
                            {
                                foreach (KeyValuePair<ListItem, object> lookupPair in lookupData.ItemLookupValues)
                                {
                                    ListItem item = lookupPair.Key;
                                    object newItemValue = null;
                                    object oldItemValue = lookupPair.Value;
                                    FieldLookupValue oldLookupValue = oldItemValue as FieldLookupValue;
                                    if (null != oldLookupValue)
                                    {
                                        int lookupId = oldLookupValue.LookupId;
                                        int newId;
                                        if (sourceProvider.m_mappingIDs.TryGetValue(lookupId, out newId) && (0 < newId))
                                        {
                                            newItemValue = new FieldLookupValue()
                                            {
                                                LookupId = newId
                                            };
                                        }
                                    }
                                    else
                                    {
                                        List<FieldLookupValue> newLookupValues = new List<FieldLookupValue>();
                                        FieldLookupValue[] oldLookupValues = oldItemValue as FieldLookupValue[];
                                        if ((null != oldLookupValues) && (0 < oldLookupValues.Length))
                                        {
                                            foreach (FieldLookupValue val in oldLookupValues)
                                            {
                                                int newId;
                                                if (sourceProvider.m_mappingIDs.TryGetValue(val.LookupId, out newId) && (0 < newId))
                                                {
                                                    newLookupValues.Add(new FieldLookupValue()
                                                    {
                                                        LookupId = newId
                                                    });
                                                }
                                            }
                                        }
                                        if (0 < newLookupValues.Count)
                                        {
                                            newItemValue = newLookupValues.ToArray();
                                        }
                                    }
                                    if (null != newItemValue)
                                    {
                                        item[lookupData.Field.InternalName] = newItemValue;
                                        item.Update();

                                        valueUpdated = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (valueUpdated)
                {
                    try
                    {
                        this.Context.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        string lookupFieldNames = string.Join(", ", m_lookups.Select(pair => pair.Value.Field.InternalName).ToArray());
                        Log.Error(ex, Constants.LOGGING_SOURCE, "Failed to set lookup values. List: '{0}', Lookup Fields: {1}.", this.List.Title, lookupFieldNames);
                    }
                }
            }
        }

        #endregion //Methods

        #region Implementation
                
        private List<Field> GetListContentSerializableFields(bool serialize)
        {
            Microsoft.SharePoint.Client.FieldCollection spfields = this.List.Fields;
            this.Context.Load(spfields);
            this.Context.ExecuteQueryRetry();

            List<Field> fields = new List<Field>();
            foreach (Field field in spfields)
            {
                if (CanFieldContentBeIncluded(field, serialize))
                {
                    fields.Add(field);
                }
            }
            return fields;
        }

        private bool CanFieldContentBeIncluded(Field field, bool serialize)
        {
            bool result = false;
            if (field.InternalName.Equals("ID", StringComparison.OrdinalIgnoreCase))
            {
                result = serialize;
            }
            else if (field.InternalName.Equals("ContentTypeId", StringComparison.OrdinalIgnoreCase) && this.List.ContentTypesEnabled)
            {
                result = true;
            }
            else if (field.InternalName.Equals("Attachments", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            else
            {
                if (!field.Hidden && !field.ReadOnlyField && (field.FieldTypeKind != FieldType.Computed))
                {
                    //Temporary disabled for custom fields
                    if (field.FieldTypeKind != FieldType.Invalid)
                    {
                        result = true;
                    }
                }
            }
            return result;
        }        

        private FieldValueProvider GetFieldValueProvider(Field field, Web web)
        {
            FieldValueProvider provider = null;

            if (null == m_dictFieldValueProviders)
            {
                m_dictFieldValueProviders = new Dictionary<string, FieldValueProvider>();
            }            
            if (!m_dictFieldValueProviders.TryGetValue(field.InternalName, out provider))
            {
                provider = new FieldValueProvider(field, web);
                m_dictFieldValueProviders.Add(field.InternalName, provider);
            }

            return provider;
        }

        private void AddIDMappingEntry(ListItem item, DataRow dataRow)
        {
            int oldId;
            string strId;
            if (dataRow.Values.TryGetValue("ID", out strId) && int.TryParse(strId, out oldId) && (0 < oldId))
            {
                if (null == m_mappingIDs)
                {
                    m_mappingIDs = new Dictionary<int, int>();
                }
                m_mappingIDs[oldId] = item.Id;

            }
        }

        private void RegisterLookupReference(FieldLookup lookupField, ListItem listitem, object itemValue)
        {
            if (null != itemValue)
            {
                if (null == m_lookups)
                {
                    m_lookups = new Dictionary<string, LookupDataRef>();
                }

                FieldLookupValue[] value = itemValue as FieldLookupValue[];

                LookupDataRef lookupRef = null;
                if (!m_lookups.TryGetValue(lookupField.InternalName, out lookupRef))
                {
                    lookupRef = new LookupDataRef(lookupField);
                    m_lookups.Add(lookupField.InternalName, lookupRef);
                }
                lookupRef.ItemLookupValues.Add(listitem, itemValue);
            }
        }

        #endregion //Implementation        
    }
}
