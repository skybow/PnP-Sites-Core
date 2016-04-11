using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.IO;
using List = Microsoft.SharePoint.Client.List;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.ListContent
{
    public class ItemPathProvider
    {
        #region Constants

        public const string FIELD_ItemDir = "FileDirRef";
        public const string FIELD_ItemName = "FileLeafRef";
        public const string FIELD_ItemType = "FSObjType";

        public const string FIELD_ItemType_FolderValue = "Folder";
        public const string FIELD_ItemType_FileValue = "File";

        #endregion //Constants

        #region Fields

        private string m_listServerRelativeUrl = "";

        #endregion //Fields

        #region Constructors

        public ItemPathProvider(List list, Web web)
        {
            var fields = list.Fields;

            m_listServerRelativeUrl = list.RootFolder.ServerRelativeUrl;                        
            this.List = list;
            this.Web = web;
        }

        #endregion //Constructors

        #region Properties

        public List List { get; private set; }
        public Web Web { get; private set; }

        public ClientRuntimeContext Context
        {
            get
            {
                return this.Web.Context;
            }
        }

        #endregion //Properties

        #region Methods

        public ListItem CreateListItem(DataRow dataRow, ProvisioningTemplate template)
        {
            ListItem listitem = null;

            string dir = null;
            string dirValue;
            if (dataRow.Values.TryGetValue(FIELD_ItemDir, out dirValue) && !string.IsNullOrEmpty(dirValue))
            {
                dir = TokenParser.CombineUrl(m_listServerRelativeUrl, dirValue);
            }

            string objType;
            if (!dataRow.Values.TryGetValue(FIELD_ItemType, out objType))
            {
                objType = "";
            }
            switch (objType)
            {
                case FIELD_ItemType_FolderValue:
                    {
                        string name;
                        if (dataRow.Values.TryGetValue(FIELD_ItemName, out name))
                        {
                            ListItemCreationInformation creationInfo = new ListItemCreationInformation()
                            {
                                UnderlyingObjectType = FileSystemObjectType.Folder,
                                LeafName = name,
                                FolderUrl = dir
                            };
                            listitem = this.List.AddItem(creationInfo);
                        }
                        break;
                    }
                case FIELD_ItemType_FileValue:
                    {
                        string name;
                        if (dataRow.Values.TryGetValue(FIELD_ItemName, out name) && !string.IsNullOrEmpty(dataRow.FileSrc))
                        {

                            if (string.IsNullOrEmpty(dir))
                            {
                                dir = this.List.RootFolder.ServerRelativeUrl;
                            }
                            FileCreationInformation creationInfo = new FileCreationInformation()
                            {
                                Overwrite = true,
                                ContentStream = template.Connector.GetFileStream(dataRow.FileSrc),
                                Url = TokenParser.CombineUrl(dir, name)
                            };
                            var newFile = this.List.RootFolder.Files.Add(creationInfo);
                            listitem = newFile.ListItemAllFields;
                        }
                            
                        break;
                    }
                default:
                    {
                        ListItemCreationInformation creationInfo = new ListItemCreationInformation()
                        {
                            FolderUrl = dir
                        };
                        listitem = this.List.AddItem(creationInfo);
                        break;
                    }
            }

            return listitem;
        }
        
        public void ExtractItemPathValues(ListItem item, Dictionary<string, string> dataRowValues, ProvisioningTemplateCreationInformation creationInfo, out string fileSrc)
        {
            fileSrc = null;

            string dir = item[FIELD_ItemDir] as string;

            var dirListRel = "";
            if (!string.IsNullOrEmpty(dir) &&
                dir.StartsWith(m_listServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                dirListRel = dir.Substring(m_listServerRelativeUrl.Length).Trim('/');
                if (!string.IsNullOrEmpty(dirListRel))
                {
                    dataRowValues[FIELD_ItemDir] = dirListRel;
                }

                if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    dataRowValues[FIELD_ItemType] = FIELD_ItemType_FolderValue;
                    dataRowValues[FIELD_ItemName] = item[FIELD_ItemName] as string;
                }
                else if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    string fileName = item[FIELD_ItemName] as string;
                    if (!string.IsNullOrEmpty(fileName) &&
                        ( this.List.BaseType == BaseType.DocumentLibrary ))
                    {
                        string fileRelUrl = TokenParser.CombineUrl(m_listServerRelativeUrl, TokenParser.CombineUrl(dirListRel, fileName)).TrimStart('/');
                        fileSrc = DownloadFile(fileRelUrl, item, creationInfo);
                        if (!string.IsNullOrEmpty(fileSrc))
                        {
                            dataRowValues[FIELD_ItemName] = fileName;
                            dataRowValues[FIELD_ItemType] = FIELD_ItemType_FileValue;
                        }
                    }
                }
            }
        }     

        #endregion //Methods

        #region Implementation

        private string DownloadFile(string fileServerRelativeURL, ListItem item, ProvisioningTemplateCreationInformation creationInfo)
        {
            string src = "";

            if (null != creationInfo.FileConnector)
            {
                ClientResult<Stream> streamResult = item.File.OpenBinaryStream();
                this.Context.ExecuteQueryRetry();
                using (Stream stream = streamResult.Value)
                {
                    creationInfo.FileConnector.SaveFileStream(fileServerRelativeURL, streamResult.Value);

                    src = Path.Combine(creationInfo.FileConnector.GetConnectionString(),
                        Path.GetDirectoryName(fileServerRelativeURL), Path.GetFileName(fileServerRelativeURL));
                }
            }

            return src;
        }

        #endregion //Implementation
    }
}
