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
        internal class ModelFileData
        {
            internal Model.File File { get; set; }
            internal string FileUrl { get; set; }
        }
        public const int FilesCountsRequestScope = 20;

        protected Web Web { get; set; }
        protected FileConnectorBase Connector { get; set; }

        private List<ModelFileData> m_files = null;

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
                m_files = new List<ModelFileData>();
            }
            m_files.Add(new ModelFileData()
            {
                File = file,
                FileUrl = fileUrl
            });
        }

        internal void UpdateFilesOverwriteFlag()
        {
            if( (null != m_files )&&( 0 < m_files.Count ) )
            {
                Web web = this.Web;
                ClientRuntimeContext ctx = web.Context;

                Dictionary<string, Microsoft.SharePoint.Client.File> dictSPFiles = new Dictionary<string, Microsoft.SharePoint.Client.File>();
                            
                int count = m_files.Count;
                int idx = 0;
                while (idx < count)
                {
                    ExceptionHandlingScope scope = new ExceptionHandlingScope(ctx);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            for (int i = idx; i < Math.Min(m_files.Count, FilesCountsRequestScope + idx); i++)
                            {
                                ModelFileData fileData = m_files[i];

                                var file = web.GetFileByServerRelativeUrl(fileData.FileUrl);
                                ctx.Load(file, f => f.Versions, f => f.Exists);

                                dictSPFiles.Add(fileData.FileUrl, file);
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
                    idx += FilesCountsRequestScope;
                }

                foreach (var fileData in m_files)
                {
                    string fileUrl = fileData.FileUrl;
                    Microsoft.SharePoint.Client.File file = null;
                    if (dictSPFiles.TryGetValue(fileUrl, out file))
                    {
                        fileData.File.Overwrite = file.Exists && file.Versions.Any();
                    }
                }
            }
        }
    }
}
