﻿using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal delegate bool ShouldProvisionTest(Web web, ProvisioningTemplate template);

    internal abstract class ObjectHandlerBase
    {
        internal bool? _willExtract;
        internal bool? _willProvision;

        private bool _reportProgress = true;
        public abstract string Name { get; }

        public bool ReportProgress
        {
            get { return _reportProgress; }
            set { _reportProgress = value; }
        }

        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        public abstract bool WillProvision(Web web, ProvisioningTemplate template);

        public abstract bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo);

        public abstract TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo);

        internal void WriteWarning(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }

        protected string Tokenize(string url, string webUrl)
        {
            if (string.IsNullOrEmpty(url))
            {
                return "";
            }
            else
            {
                if (url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    return url.Substring(url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/theme", "{themecatalog}");
                }
                if (url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    return url.Substring(url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/masterpage", "{masterpagecatalog}");
                }
                if (url.IndexOf(webUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    return url.Replace(webUrl, "{site}");
                }
                if (url.IndexOf(webUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    return url.Substring(url.IndexOf(webUrl, StringComparison.InvariantCultureIgnoreCase)).Replace(webUrl, "{site}");
                }
                else
                {
                    Uri r;
                    if (Uri.TryCreate(webUrl, UriKind.Absolute, out r))
                    {
                        var webUrlPathAndQuery = HttpUtility.UrlDecode(r.PathAndQuery);
                        if (webUrlPathAndQuery == "/") {
                            if (url.StartsWith("/"))
                            {
                                var urlParts = url.Split('/');
                                urlParts[0] = "{site}";
                                return string.Join("/", urlParts);
                            }
                            else
                            {
                                Uri r2;
                                if (Uri.TryCreate(url, UriKind.Absolute, out r2))
                                {
                                    return "{site}" + r2.PathAndQuery.Substring(1);
                                }
                            }
                            return url;
                        } 
                        else if (url.IndexOf(webUrlPathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                        {
                            return url.Replace(webUrlPathAndQuery, "{site}");
                        }
                    }
                }

                // nothing to tokenize...
                return url;
            }
        }
    }
}
