using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Feature = OfficeDevPnP.Core.Framework.Provisioning.Model.Feature;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectFeatures : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Features"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                //----------------------------
                //Feature activation order
                //  - Site Collection Features
                //  - Web Features
                //----------------------------
                //Feature Deactivation order
                // - Web features
                // - Site Collection Features
                //----------------------------

                // if this is a sub site then we're not enabling the site collection scoped features
                if (!web.IsSubSite())
                {
                    var siteFeaturesToActivate = template.Features.SiteFeatures.AsQueryable().Where(f => !f.Deactivate).ToList();
                    ProvisionFeaturesImplementation<Site>(context.Site, siteFeaturesToActivate, scope);
                }

                var webFeaturesToActivate = template.Features.WebFeatures.AsQueryable().Where(f => !f.Deactivate).ToList();
                ProvisionFeaturesImplementation<Web>(web, webFeaturesToActivate, scope);

                var webFeaturesToDeactivate = template.Features.WebFeatures.AsQueryable().Where(f => f.Deactivate).ToList();
                ProvisionFeaturesImplementation<Web>(web, webFeaturesToDeactivate, scope);

                if (!web.IsSubSite())
                {
                    var siteFeaturesToDeactivate = template.Features.SiteFeatures.AsQueryable().Where(f => f.Deactivate).ToList();
                    ProvisionFeaturesImplementation<Site>(context.Site, siteFeaturesToDeactivate, scope);
                }


            }
            return parser;
        }

        private static void ProvisionFeaturesImplementation<T>(T parent, IEnumerable<Feature> features, PnPMonitoredScope scope)
        {
            var activeFeatures = new List<Microsoft.SharePoint.Client.Feature>();
            Web web = null;
            Site site = null;
            if (parent is Site)
            {
                site = parent as Site;
                site.Context.Load(site.Features, fs => fs.Include(f => f.DefinitionId));
                site.Context.ExecuteQueryRetry();
                activeFeatures = site.Features.ToList();
            }
            else
            {
                web = parent as Web;
                web.Context.Load(web.Features, fs => fs.Include(f => f.DefinitionId));
                web.Context.ExecuteQueryRetry();
                activeFeatures = web.Features.ToList();
            }

            if (features != null)
            {
                bool needToRetry = false;
                do
                {
                    needToRetry = false;
                    var unProvisionFeatures = new List<Feature>();

                    foreach (var feature in features)
                    {
                        try
                        {

                            if (!feature.Deactivate)
                            {
                                if (activeFeatures.FirstOrDefault(f => f.DefinitionId == feature.Id) == null)
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Features_Activating__0__scoped_feature__1_, site != null ? "site" : "web", feature.Id);
                                    if (site != null)
                                    {
                                        site.ActivateFeature(feature.Id);
                                    }
                                    else
                                    {
                                        web.ActivateFeature(feature.Id);
                                    }
                                }
                            }
                            else
                            {
                                if (activeFeatures.FirstOrDefault(f => f.DefinitionId == feature.Id) != null)
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Features_Deactivating__0__scoped_feature__1_, site != null ? "site" : "web", feature.Id);
                                    if (site != null)
                                    {
                                        site.DeactivateFeature(feature.Id);
                                    }
                                    else
                                    {
                                        web.DeactivateFeature(feature.Id);
                                    }

                                    needToRetry = true;
                                }
                            }
                        }
                        catch (ServerException ex)
                        {
                            if (ex.ServerErrorTypeName == "Microsoft.SharePoint.SPFeatureDependencyNotActivatedException")
                            {
                                string message = ex.Message;

                                string strGuidRegex = @"\b[A-F0-9]{8}(?:-[A-F0-9]{4}){3}-[A-F0-9]{12}\b";
                                Regex guidRegex = new Regex(strGuidRegex, RegexOptions.IgnoreCase);
                                var matches = guidRegex.Matches(message).Cast<Match>().Select(m => m.Value).Where(x => !x.Equals(feature.Id.ToString(), StringComparison.CurrentCultureIgnoreCase));
                                InsertFeaturetoCorrectProvisionOrder(unProvisionFeatures, feature, matches);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(Constants.LOGGING_SOURCE_FRAMEWORK_PROVISIONING, " Error in ProvisionFeatures: {0} - {1}", ex.Message, ex.StackTrace);
                        }
                    }

                    features = unProvisionFeatures;
                } while (needToRetry);
            }
        }


        private static void InsertFeaturetoCorrectProvisionOrder(List<Feature> features, Feature featureToInsert, IEnumerable<string> dependenceFeaturesIds)
        {
            if (featureToInsert.Deactivate)
            {
                int indexToInsertForDeactivation = features.FindLastIndex((f) =>
                {
                    return dependenceFeaturesIds.Contains(f.Id.ToString());
                });

                if (indexToInsertForDeactivation == -1)
                {
                    features.Add(featureToInsert);
                }
                else
                {
                    features.Insert(indexToInsertForDeactivation, featureToInsert);
                }
            }
            else
            {
                int indexToInsertForActivation = features.FindLastIndex((f) =>
                {
                    return dependenceFeaturesIds.Contains(f.Id.ToString());
                });

                features.Insert(indexToInsertForActivation + 1, featureToInsert);
            }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                bool isSubSite = web.IsSubSite();
                var webFeatures = web.Features;
                var siteFeatures = context.Site.Features;

                context.Load(webFeatures, fs => fs.Include(f => f.DefinitionId));
                if (!isSubSite)
                {
                    context.Load(siteFeatures, fs => fs.Include(f => f.DefinitionId));
                }
                context.ExecuteQueryRetry();

                var features = new Features();
                foreach (var feature in webFeatures)
                {
                    features.WebFeatures.Add(new Feature() { Deactivate = false, Id = feature.DefinitionId });
                }

                // if this is a sub site then we're not creating  site collection scoped feature entities
                if (!isSubSite)
                {
                    foreach (var feature in siteFeatures)
                    {
                        features.SiteFeatures.Add(new Feature() { Deactivate = false, Id = feature.DefinitionId });
                    }
                }

                template.Features = features;

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate, isSubSite);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate, bool isSubSite)
        {
            List<Guid> featuresToExclude = new List<Guid>();
            // Seems to be an feature left over on some older online sites...
            featuresToExclude.Add(Guid.Parse("d70044a4-9f71-4a3f-9998-e7238c11ce1a"));

            if (!isSubSite)
            {
                foreach (var feature in baseTemplate.Features.SiteFeatures)
                {
                    int index = template.Features.SiteFeatures.FindIndex(f => f.Id.Equals(feature.Id));

                    if (index > -1)
                    {
                        template.Features.SiteFeatures.RemoveAt(index);
                    }
                }

                foreach (var feature in featuresToExclude)
                {
                    int index = template.Features.SiteFeatures.FindIndex(f => f.Id.Equals(feature));

                    if (index > -1)
                    {
                        template.Features.SiteFeatures.RemoveAt(index);
                    }
                }

            }

            foreach (var feature in baseTemplate.Features.WebFeatures)
            {
                int index = template.Features.WebFeatures.FindIndex(f => f.Id.Equals(feature.Id));

                if (index > -1)
                {
                    template.Features.WebFeatures.RemoveAt(index);
                }
            }

            foreach (var feature in featuresToExclude)
            {
                int index = template.Features.WebFeatures.FindIndex(f => f.Id.Equals(feature));

                if (index > -1)
                {
                    template.Features.WebFeatures.RemoveAt(index);
                }
            }

            return template;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Features.SiteFeatures.Any() || template.Features.WebFeatures.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }
}
