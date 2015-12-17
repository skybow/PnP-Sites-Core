using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    public class IdToken : TokenDefinition
    {
        protected string NewId { get; private set; }
        public IdToken(Web web, string newId, string oldId)
            : base(web, oldId)
        {
            this.NewId = newId;
        }

        public override string GetReplaceValue()
        {
            return NewId;
        }
    }
}
