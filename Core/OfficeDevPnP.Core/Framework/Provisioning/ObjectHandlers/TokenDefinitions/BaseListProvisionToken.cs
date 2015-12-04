using System;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    public abstract class BaseListProvisionToken : TokenDefinition
    {
        protected Guid _listId;

        protected BaseListProvisionToken(Web web, Guid currentListId, string token) : base(web,token)
        {
             _listId = currentListId;
        }
    }
}
