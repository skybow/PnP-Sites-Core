using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermGroupNameToken : TokenDefinition
    {
        public SiteCollectionTermGroupNameToken(Web web)
            : base(web, "~sitecollectiontermgroupname", "{sitecollectiontermgroupname}")
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

                    CacheValue = termGroup.Name.ToString();
                }
                catch (Exception ex)
                {

                }
            }
            return CacheValue;
        }
    }
}