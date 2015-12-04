using System;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ListNameProvisionToken : BaseListProvisionToken
    {
        public ListNameProvisionToken(Web web, Guid currentListId, Guid templateistId)
            : base(web, currentListId, templateistId.ToString("B").ToUpper())
        {
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _listId.ToString("B").ToUpper();
            }
            return CacheValue;
        }
    }
}
