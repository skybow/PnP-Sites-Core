using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class KeywordsTermStoreIdToken : TokenDefinition
    {
        public KeywordsTermStoreIdToken(Web web)
            : base(web, "~keywordstermstoreid", "{keywordstermstoreid}")
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
                    var termStore = session.GetDefaultKeywordsTermStore();
                    Web.Context.Load(termStore, t => t.Id);
                    Web.Context.ExecuteQueryRetry();
                    CacheValue = termStore.Id.ToString();
                }
                catch( Exception ex )
                {
                    Log.Error(ex, Constants.LOGGING_SOURCE, 
                        "Failed to retrive {0} token value.", String.Join( ", ", GetTokens() ));
                }
            }
            return CacheValue;
        }
    }
}