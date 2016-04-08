using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class AssociatedGroupToken : TokenDefinition
    {
        private enum AssociatedGroupType
        {
            Owners,
            Members,
            Visitors
        }

        private class AssociatedGroupTokenLoader
        {
            public Web Web { get; private set; }
            private Dictionary<AssociatedGroupType, string> _dictValues = null;

            public AssociatedGroupTokenLoader(Web web)
            {
                this.Web = web;
            }

            public string GetAssosiatedGroupTitle(AssociatedGroupType groupType)
            {
                EnsureLoadAssotiatedGroups();

                string groupTitle;
                if (!this._dictValues.TryGetValue(groupType, out groupTitle))
                {
                    groupTitle = "";
                }
                return groupTitle;
            }

            private void EnsureLoadAssotiatedGroups()
            {
                if (null == _dictValues)
                {
                    _dictValues = new Dictionary<AssociatedGroupType, string>();

                    try
                    {
                        var context = this.Web.Context as ClientContext;
                        context.Load(Web, w => w.AssociatedOwnerGroup.Title, w => w.AssociatedMemberGroup.Title, w => w.AssociatedVisitorGroup.Title);
                        context.ExecuteQueryRetry();

                        if(null != this.Web.AssociatedOwnerGroup)
                        {
                            this._dictValues[AssociatedGroupType.Owners] =  Web.AssociatedOwnerGroup.Title;
                        }
                        if(null != this.Web.AssociatedMemberGroup)
                        {
                            this._dictValues[AssociatedGroupType.Members] =  Web.AssociatedMemberGroup.Title;
                        }
                        if(null != this.Web.AssociatedVisitorGroup)
                        {
                            this._dictValues[AssociatedGroupType.Visitors] =  Web.AssociatedVisitorGroup.Title;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error( ex, Constants.LOGGING_SOURCE, "Failed to load web associated groups." );
                    }
                }
            }
        }

        private AssociatedGroupType _groupType;

        private AssociatedGroupTokenLoader _loader = null;

        private AssociatedGroupToken(Web web, AssociatedGroupType groupType, AssociatedGroupTokenLoader loader):
            base(web, string.Format("{{associated{0}group}}", groupType.ToString().TrimEnd('s') ))
        {
            _groupType = groupType;
            _loader = loader;
        }

        public override string GetReplaceValue()
        {            
            if (string.IsNullOrEmpty(CacheValue))
            {
                this.CacheValue = _loader.GetAssosiatedGroupTitle(this._groupType);                
            }
            return CacheValue;
        }

        public static AssociatedGroupToken[] CreateAssociatedGroupsTokens(Web web)
        {
            AssociatedGroupTokenLoader loader = new AssociatedGroupTokenLoader(web);

            AssociatedGroupToken[] tokens = new AssociatedGroupToken[]
            {
                new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.Owners, loader),
                new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.Members, loader),
                new AssociatedGroupToken(web, AssociatedGroupToken.AssociatedGroupType.Visitors, loader)
            };
            return tokens;
        }        
    }
}