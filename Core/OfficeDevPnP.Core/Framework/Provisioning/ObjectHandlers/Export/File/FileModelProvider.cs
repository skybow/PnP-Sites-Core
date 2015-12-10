using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.File
{
    internal class FileModelProvider
    {
        protected Web Web { get; set; }
        protected FileConnectorBase Connector { get; set; }
        private readonly Regex webUrlReplceReqEx;
        private readonly Regex webIdReplceReqEx;

        public FileModelProvider(Web web, FileConnectorBase connector)
        {
            Web = web;
            Connector = connector;
            webUrlReplceReqEx = new Regex(web.ServerRelativeUrl, RegexOptions.Singleline);
            webIdReplceReqEx = new Regex(web.Id.ToString(), RegexOptions.Singleline);
        }

        public Model.File GetFile(string pageUrl)
        {
            var webPartsXml = Web.GetWebPartsXml(pageUrl);

            webPartsXml = this.TokenizeXml(webPartsXml);

            var provider = new WebPartsModelProvider(Web, pageUrl);
            var webPartsModels = provider.Retrieve(webPartsXml, pageUrl);

            var needToOverride = this.NeedToOverrideFile(Web, pageUrl);

            var folderPath = this.GetFolderPath(pageUrl);

            var localFilePath = this.GetFilePath(pageUrl);

            return new Model.File(localFilePath, folderPath, needToOverride, webPartsModels, null);
        }

        private string GetFolderPath(string pageUrl)
        {
            var folder = webUrlReplceReqEx.Replace(pageUrl, "~site");
            return folder.Substring(0, folder.LastIndexOf("/", StringComparison.Ordinal));
        }

        private string GetFilePath(string pageUrl)
        {
            var fileName = Path.GetFileName(pageUrl);
            var filePath = Path.Combine(Path.GetDirectoryName(pageUrl), fileName).TrimStart('\\');

            return Path.Combine(this.Connector.GetConnectionString(), filePath);
        }

        private string TokenizeXml(string xml)
        {
            xml = webUrlReplceReqEx.Replace(xml, "~site");
            return webIdReplceReqEx.Replace(xml, "~siteid");
        }

        private bool NeedToOverrideFile(Web web, string pageUrl)
        {
            var file = web.GetFileByServerRelativeUrl(pageUrl);
            web.Context.Load(file, f => f.Versions);
            web.Context.ExecuteQueryRetry();
            return file.Versions.Any();
        }
    }
}
