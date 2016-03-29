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

        private string m_listServerRelativeUrl = null;

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
                dir = TokenParser.CombineUrl(this.Web.ServerRelativeUrl, dirValue);
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
                        if (dataRow.Values.TryGetValue(FIELD_ItemName, out name) && !string.IsNullOrEmpty( dataRow.FileSrc ) )
                        {
                            using (Stream stream = template.Connector.GetFileStream(dataRow.FileSrc))
                            {
                                if (string.IsNullOrEmpty(dir))
                                {
                                    dir = this.List.RootFolder.ServerRelativeUrl;
                                }
                                FileCreationInformation creationInfo = new FileCreationInformation()
                                {
                                    Overwrite = true,
                                    ContentStream = stream,
                                    Url = TokenParser.CombineUrl(dir, name)
                                };
                                var newFile = this.List.RootFolder.Files.Add(creationInfo);
                                this.Context.Load(newFile);
                                this.Context.ExecuteQueryRetry();

                                listitem = newFile.ListItemAllFields;
                            }
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

        public void CheckInIOfNeeded(ListItem listitem)
        {
            if ((listitem.FileSystemObjectType == FileSystemObjectType.File) &&
                        (null != listitem.File.ServerObjectIsNull) && (!(bool)listitem.File.ServerObjectIsNull) &&
                        (listitem.File.CheckOutType != CheckOutType.None))
            {
                listitem.File.CheckIn("", CheckinType.MajorCheckIn);
                this.Context.ExecuteQueryRetry();
            }
        }

        public void ExtractItemPathValues(ListItem item, Dictionary<string, string> dataRowValues, ProvisioningTemplateCreationInformation creationInfo, out string fileSrc)
        {
            fileSrc = null;

            string dir = item[FIELD_ItemDir] as string;
            string dirWebRel = "";
            if (!string.IsNullOrEmpty(dir) &&
                dir.StartsWith(this.Web.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                dirWebRel = dir.Substring(this.Web.ServerRelativeUrl.Length).TrimStart('/');
            }
            if (!string.IsNullOrEmpty(dirWebRel))
            {
                if (!dir.Equals(m_listServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
                {
                    dataRowValues[FIELD_ItemDir] = dirWebRel;
                }

                if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    dataRowValues[FIELD_ItemType] = FIELD_ItemType_FolderValue;
                    dataRowValues[FIELD_ItemName] = item[FIELD_ItemName] as string;
                }
                else if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    string fileName = item[FIELD_ItemName] as string;
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        string fileRelUrl = TokenParser.CombineUrl(dirWebRel, fileName);
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

        private string DownloadFile(string fileWebRelURL, ListItem item, ProvisioningTemplateCreationInformation creationInfo)
        {
            string src = "";

            if (null != creationInfo.FileConnector)
            {
                this.Context.Load(item.File);
                this.Context.ExecuteQueryRetry();
                if ((null != item.File.ServerObjectIsNull) &&
                    !(bool)item.File.ServerObjectIsNull)
                {
                    ClientResult<Stream> streamResult = item.File.OpenBinaryStream();
                    this.Context.ExecuteQueryRetry();
                    using (Stream stream = streamResult.Value)
                    {
                        creationInfo.FileConnector.SaveFileStream(fileWebRelURL, streamResult.Value);

                        src = Path.Combine(creationInfo.FileConnector.GetConnectionString(),
                            Path.GetDirectoryName(fileWebRelURL), Path.GetFileName(fileWebRelURL));
                    }
                }
            }

            return src;
        }

        /*
        private Folder EnsureFolder( string urlWebRelative  )
        {
            Folder currentFolder = m_list.RootFolder;
            string rootUrl = currentFolder.ServerRelativeUrl;            

            string targetFolderServerUrl = TokenParser.CombineUrl( this.m_web.ServerRelativeUrl, urlWebRelative );

            // Get remaining parts of the path and split

            var folderRootRelativeUrl = targetFolderServerUrl.Substring(currentFolder.ServerRelativeUrl.Length);
            var childFolderNames = folderRootRelativeUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentCount = 0;

            foreach (var folderName in childFolderNames)
            {
                currentCount++;

                // Find next part of the path
                var folderCollection = currentFolder.Folders;
                folderCollection.Context.Load(folderCollection);
                folderCollection.Context.ExecuteQueryRetry();
                Folder nextFolder = null;
                foreach (Folder existingFolder in folderCollection)
                {
                    if (string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        nextFolder = existingFolder;
                        break;
                    }
                }

                // Or create it
                if (nextFolder == null)
                {
                    ListItem itemFolder = this.m_list.AddItem(new ListItemCreationInformation()
                    {
                        UnderlyingObjectType = FileSystemObjectType.Folder,
                        LeafName = folderName,
                        FolderUrl = currentFolder.ServerRelativeUrl
                    });
                    this.m_web.Context.Load(itemFolder, i => i.Folder);                    
                    this.m_web.Context.ExecuteQueryRetry();
                    nextFolder = itemFolder.Folder;
                }

                currentFolder = nextFolder;
            }

            return currentFolder;
        }*/

        #endregion //Implementation
    }
}
