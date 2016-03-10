using System;
using System.IO;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.File
{
    internal class FileModelProvider
    {
        protected Web Web { get; set; }
        protected FileConnectorBase Connector { get; set; }

        private Dictionary<string, Model.File> m_files = null;

        public FileModelProvider(Web web, FileConnectorBase connector)
        {
            Web = web;
            Connector = connector;
        }

        public Model.File GetFile(string pageUrl)
        {
            Model.File file = null;
            if (pageUrl.StartsWith(Web.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                var provider = new WebPartsModelProvider(Web);
                var webPartsModels = provider.Retrieve(pageUrl);

                var folderPath = this.GetFolderPath(pageUrl);

                var localFilePath = this.GetFilePath(pageUrl);

                file = new Model.File(localFilePath, folderPath, false, webPartsModels, null);

                AddFileToSetOverrideFlagStack(file, pageUrl);
            }
            return file;
        }

        private string GetFolderPath(string pageUrl)
        {
            var folder = "";
            if (pageUrl.StartsWith(Web.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                folder = TokenParser.CombineUrl("~site", pageUrl.Substring(Web.ServerRelativeUrl.Length));
            }
            return folder.Substring(0, folder.LastIndexOf("/", StringComparison.Ordinal));
        }

        private string GetFilePath(string pageUrl)
        {
            var fileName = Path.GetFileName(pageUrl);
            var filePath = Path.Combine(Path.GetDirectoryName(pageUrl), fileName).TrimStart('\\');

            return Path.Combine(this.Connector.GetConnectionString(), filePath);
        }

        private void AddFileToSetOverrideFlagStack( Model.File file, string fileUrl )
        {
            if (null == m_files)
            {
                m_files = new Dictionary<string, Model.File>();
            }
            m_files.Add( fileUrl, file );
        }

        internal void UpdateFilesOverwriteFlag()
        {
            if( (null != m_files )&&( 0 < m_files.Count ) )
            {
                Web web = this.Web;
                ClientRuntimeContext ctx = web.Context;

                Dictionary<string, Microsoft.SharePoint.Client.File> dictSPFiles = new Dictionary<string, Microsoft.SharePoint.Client.File>();
                            
                ExceptionHandlingScope scope = new ExceptionHandlingScope(ctx);

                using (scope.StartScope())
                {
                    using (scope.StartTry())
                    {
                        foreach (KeyValuePair<string, Model.File> pair in m_files)
                        {
                            string fileUrl = pair.Key;

                            var file = web.GetFileByServerRelativeUrl(fileUrl);
                            web.Context.Load(file, f => f.Versions, f => f.Exists);

                            dictSPFiles.Add(fileUrl, file);
                        }
                    }
                    using (scope.StartCatch())
                    {
                    }
                    using (scope.StartFinally())
                    {
                    }
                }
                ctx.ExecuteQuery();

                foreach (KeyValuePair<string, Microsoft.SharePoint.Client.File> pair in dictSPFiles)
                {
                    Microsoft.SharePoint.Client.File spfile = pair.Value;
                    if ((null != spfile) && spfile.Exists)
                    {
                        Model.File file = m_files[pair.Key];
                        file.Overwrite = pair.Value.Versions.Any();
                    }
                }
            }
        }
    }
}
