using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermGroupIdToken : TokenDefinition
    {
        public SiteCollectionTermGroupIdToken(Web web)
            : base(web, "~sitecollectiontermgroupid", "{sitecollectiontermgroupid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (!ValueRetrieved)
            {
                ValueRetrieved = true;
                try
                {
                    // The token is requested. Check if the group exists and if not, create it
                    var site = (Web.Context as ClientContext).Site;
                    var session = TaxonomySession.GetTaxonomySession(site.Context);
                    var termstore = session.GetDefaultSiteCollectionTermStore();
                    var termGroup = termstore.GetSiteCollectionGroup(site, true);
                    site.Context.Load(termGroup);
                    site.Context.ExecuteQueryRetry();

                    CacheValue = termGroup.Id.ToString();
                }
                catch (Exception)
                {
                }
            }
            return CacheValue;
        }
    }
}