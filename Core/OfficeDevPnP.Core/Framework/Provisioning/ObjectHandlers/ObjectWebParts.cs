using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebParts : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Parts"; }
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Lists.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var collList = web.Lists;
                var lists = web.Context.LoadQuery(collList);

                web.Context.ExecuteQueryRetry();

                _willExtract = lists.Any();
            }
            return _willExtract.Value;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            web.EnsureProperty(x => x.Id);
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                foreach (var listPage in template.ListPages)
                {
                    listPage.PageUrl = parser.ParseString(listPage.PageUrl);
                    foreach (var webPartEntity in listPage.WebPartEntities)
                    {
                        try
                        {
                            webPartEntity.WebPartXml = parser.ParseString(webPartEntity.WebPartXml, Guid.Empty.ToString());
                            web.AddWebPartToWebPartPage(listPage.PageUrl, webPartEntity);
                        }
                        catch (Exception exception)
                        {
                            var message = string.Format("Cannot export web parts for: {0}", listPage.PageUrl);
                            scope.LogError(exception, message);
                        }
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            web.EnsureProperties(x => x.Url, x => x.Id);
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                try
                {
                    var forms = template.Lists.SelectMany(x => x.Forms).ToArray();
                    var lispPages = forms.Select(x => GetListPage(web, x.ServerRelativeUrl, scope)).ToList();
                    lispPages.AddRange(this.GetListPagesFromViews(web, template, scope));
                    template.ListPages = lispPages;
                }
                catch (Exception exception)
                {
                    scope.LogError(exception, "Extract web parts error");
                }
            }
            return template;
        }

        private IEnumerable<ListPage> GetListPagesFromViews(Web web, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            var result = new List<ListPage>();
            var views = template.Lists.SelectMany(x => x.Views).ToArray();
            var viewIdReqEx = new Regex(@"Name=""(.*?)""");
            foreach (var view in views)
            {
                var viewId = viewIdReqEx.Match(view.SchemaXml).Value;
                var page = GetListPage(web, view.PageUrl, scope);
                page.WebPartEntities = GetWebPartsWithoutDefaultView(viewId, page.WebPartEntities).ToList();
                if (page.WebPartEntities.Any())
                {
                    result.Add(page);
                }
            }
            return result;
        }

        private IEnumerable<WebPartEntity> GetWebPartsWithoutDefaultView(string viewId, IEnumerable<WebPartEntity> entities)
        {
            return entities.Where(x => x.WebPartXml.IndexOf(viewId, StringComparison.Ordinal) == -1);
        }

        private ListPage GetListPage(Web web, string pageUrl, PnPMonitoredScope scope)
        {
            var webPartsXml = this.ExtractWebPartsXml(web, pageUrl, scope);
            var webUrlReplceReqEx = new Regex(web.ServerRelativeUrl, RegexOptions.Singleline);
            var webIdReplceReqEx = new Regex(web.Id.ToString(), RegexOptions.Singleline);
            webPartsXml = webUrlReplceReqEx.Replace(webPartsXml, "~site");
            webPartsXml = webIdReplceReqEx.Replace(webPartsXml, "~siteid");
            var provider = new WebPartsEntityProvider(web, pageUrl);
            var entities = provider.Retrieve(webPartsXml, pageUrl);
            return new ListPage
            {
                PageUrl = webUrlReplceReqEx.Replace(pageUrl, "~site"),
                WebPartEntities = entities
            };
        }

        private string ExtractWebPartsXml(Web web, string pageUrl, PnPMonitoredScope scope)
        {
            try
            {
                var url = string.Format("{0}/_vti_bin/Webpartpages.asmx", web.Url);
                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(url);

                endpointRequest.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested;

                endpointRequest.Method = "POST";
                endpointRequest.Accept = "text/xml; charset=utf-8";
                endpointRequest.ContentType = "text/xml; charset=utf-8";
                endpointRequest.UseDefaultCredentials = false;
                endpointRequest.Credentials = web.Context.Credentials;
                var message = "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"
                    + "<soap:Body>"
                    + "   <GetWebPartProperties2 xmlns=\"http://microsoft.com/sharepoint/webpartpages\">"
                    + "      <pageUrl>" + pageUrl + "</pageUrl>"
                    + "      <storage>Shared</storage>"
                    + "      <behavior>Version3</behavior>"
                    + "   </GetWebPartProperties2>"
                    + "</soap:Body>"
                    + "</soap:Envelope>";

                byte[] bytes = Encoding.UTF8.GetBytes(message);
                endpointRequest.ContentLength = bytes.Length;
                Stream dataStream = endpointRequest.GetRequestStream();
                dataStream.Write(bytes, 0, bytes.Length);
                dataStream.Close();
                string webPartsSchemas;
                using (HttpWebResponse response = endpointRequest.GetResponse() as HttpWebResponse)
                {
                    StreamReader reader = new StreamReader(response.GetResponseStream());
                    webPartsSchemas = reader.ReadToEnd();
                }
                return webPartsSchemas;
            }
            catch (Exception exception)
            {
                scope.LogError(exception, "Web Parts Service call error");
                return string.Empty;
            }
        }
    }
}
