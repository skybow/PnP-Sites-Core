using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermStoreIdToken : TokenDefinition
    {
        public SiteCollectionTermStoreIdToken(Web web)
            : base(web, "~sitecollectiontermstoreid", "{sitecollectiontermstoreid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (!ValueRetrieved)
            {
                ValueRetrieved = true;
                try
                {
                    TaxonomySession session = TaxonomySession.GetTaxonomySession(Web.Context);
                    var termStore = session.GetDefaultSiteCollectionTermStore();
                    Web.Context.Load(termStore, t => t.Id);
                    Web.Context.ExecuteQueryRetry();
                    if (termStore != null)
                    {
                        CacheValue = termStore.Id.ToString();
                    }
                }
                catch (System.Exception)
                {

                }
            }
            return CacheValue;
        }
    }
}