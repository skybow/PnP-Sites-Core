using System;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ListIdProvisionToken : BaseListProvisionToken
    {
        public ListIdProvisionToken(Web web, Guid currentListId, Guid templateistId) 
            : base(web, currentListId,templateistId.ToString("D"))
        {
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _listId.ToString("D");
            }
            return CacheValue;
        }
    }
}
